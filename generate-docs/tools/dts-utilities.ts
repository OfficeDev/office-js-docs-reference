import * as ts from "typescript";

// capturing these because of eccentricities with the compiler ordering
let topClass: ClassStruct = null;
let lastItem: ClassStruct | FieldStruct = null;

enum ClassType {
    Class = "Class",
    Interface = "Interface",
    Enum = "Enum",
}

enum FieldType {
    Property = "Property",
    Method = "Method",
    Event = "Event",
    Enum = "Enum",
}

class FieldStruct {
    public declarationString: string;
    public comment: string;
    public type: FieldType;
    public name: string;

    constructor(decString: string, commentString: string, fieldType: FieldType, fieldName: string) {
        this.declarationString = decString;
        this.comment = commentString;
        this.type = fieldType;
        this.name = fieldName;
    }
}

class ClassStruct {
    public declarationString: string;
    public comment: string;
    public type: ClassType;
    public fields: FieldStruct[];

    constructor(decString, commentString: string, classType: ClassType) {
        this.declarationString = decString;
        this.comment = commentString;
        this.type = classType;
        this.fields = [];
    }

    public copyWithoutFields(): ClassStruct {
        return new ClassStruct(this.declarationString, this.comment, this.type);
    }

    public sortFields(): void {
        this.fields.sort((a, b) => {
            if (a.declarationString === b.declarationString) {
                return 0;
            } else {
                return a.declarationString.replace("readonly ", "") < b.declarationString.replace("readonly ", "") ? -1 : 1;
            }
        });
    }

    public getClassName(): string {
        return this.declarationString.substring(this.declarationString.lastIndexOf(" ") + 1);
    }
}

export class APISet {
    public api: ClassStruct[];
    constructor() {
        this.api = [];
    }

    public addClass(clas: ClassStruct): void {
        this.api.push(clas);
    }

    public containsClass(clas: ClassStruct): boolean {
        let found: boolean = false;
        this.api.forEach((element) => {
            if (element.declarationString === clas.declarationString) {
                found = true;
            }
        });

        return found;
    }

    public containsField(clas: ClassStruct, field: FieldStruct): boolean {
        let found: boolean = false;
        this.api.forEach((element) => {
            if (element.declarationString === clas.declarationString) {
                element.fields.forEach((thisField) => {
                    if (thisField.declarationString === field.declarationString) {
                        found = true;
                    }
                });
            }
        });

        return found;
    }

    // finds the new fields and classes
    public diff(other: APISet): APISet {
        const diffAPI: APISet = new APISet();
        this.api.forEach((element) => {
            if (other.containsClass(element)) {
                let classShell: ClassStruct = null;
                element.fields.forEach((field) => {
                    if (!other.containsField(element, field)) {
                        if (classShell === null) {
                            classShell = element.copyWithoutFields();
                            diffAPI.addClass(classShell);
                        }

                        classShell.fields.push(field);
                    }
                });
            } else {
                diffAPI.addClass(element);
            }
        });

        return diffAPI;
    }

    public getAsDTS(): string {
        this.sort();
        const output: string[] = [];
        this.api.forEach((clas) => {
            output.push(clas.comment.trim());
            output.push(clas.declarationString + " {");
            clas.fields.forEach((field) => {
                output.push("    " + field.comment);
                if (field.type === FieldType.Enum) {
                    output.push("    " + field.declarationString + ",");
                } else {
                    output.push("    " + field.declarationString);
                }
            });
            output.push("}");
        });
        return output.join("\n");
    }

    public getAsMarkdown(relativePath: string): string {
        this.sort();
        // table header
        let output: string = "| Class | Fields | Description |\n|:---|:---|:---|\n";
        this.api.forEach((clas) => {
            // Ignore the following:
            // - Enums.
            // - LoadOptions interfaces
            // - *Data classes for set/load methods
            if (clas.type !== ClassType.Enum &&
                !clas.getClassName().endsWith("LoadOptions") &&
                !clas.getClassName().endsWith("Data")) {
                const className = clas.getClassName();
                output += "|[" + className + "](/"
                    + relativePath + className.toLowerCase() + ")|";
                let first: boolean = true;
                clas.fields.forEach((field) => {
                    // Ignore the following:
                    // - String literal overloads.
                    // - `load`, `set`, `track`, `untrack`, and `toJSON` methods
                    // - The `context` property.
                    // - Static fields.
                    if (field.declarationString.search(/([a-zA-Z]+)\??: (\"[a-zA-Z]*\").*:/g) < 0 &&
                        field.name !== "load" &&
                        field.name !== "set" &&
                        field.name !== "toJSON" &&
                        field.name !== "context" &&
                        field.name !== "track" &&
                        field.name !== "untrack" &&
                        !field.declarationString.includes("static ")) {
                        if (first) {
                            first = false;
                        } else {
                            output += "||";
                        }

                        // remove unnecessary parts of the declaration string
                        let newItemText = field.declarationString.replace(/;/g, "");
                        newItemText = newItemText.substring(0, newItemText.lastIndexOf(":")).replace("readonly ", "");
                        newItemText = newItemText.replace(/\|/g, "\\|").replace(/\n|\t/gm, "");
                        if (field.type === FieldType.Property) {
                            newItemText = newItemText.replace(/\?/g, "");
                        } 
                        
                        newItemText = newItemText.replace(/\<any\>/g, "");
                        

                        let tableLine = "[" + newItemText + "]("
                            + buildFieldLink(relativePath, className, field) + ")|";
                        tableLine += extractFirstSentenceFromComment(field.comment);
                        output += tableLine + "|\n";
                    }
                });
            }
        });
        return output;
    }

    public sort(): void {
        this.api.forEach((element) => {
            element.sortFields();
        });

        this.api.sort((a, b) => {
            if (a.getClassName() === b.getClassName()) {
                return 0;
            } else {
                return a.getClassName() < b.getClassName() ? -1 : 1;
            }
        });
    }
}

function extractFirstSentenceFromComment(commentText) {
    const firstSentenceIndex = commentText.indexOf("* ") + 2;
    const multiSentenceEndIndex = commentText.indexOf(". ", firstSentenceIndex);
    const lineBreakEndIndex = commentText.indexOf("\n", firstSentenceIndex);
    const singleLineEndIndex = commentText.indexOf("\*/", firstSentenceIndex);

    let endIndex;
    if (multiSentenceEndIndex > 0 && lineBreakEndIndex > 0) {
        endIndex = Math.min(multiSentenceEndIndex + 1, lineBreakEndIndex);
    } else if (multiSentenceEndIndex === -1 && lineBreakEndIndex === -1) {
        endIndex = singleLineEndIndex;
    } else {
        endIndex = Math.max(multiSentenceEndIndex + 1, lineBreakEndIndex);
    }

    return commentText.substring(firstSentenceIndex, endIndex).trim();
}

function buildFieldLink(relativePath: string, className: string, field: FieldStruct) {
    let fieldLink: string;
    if (field.type === FieldType.Method) {
        // Remove anonymous types before proceeding.
        let fieldString = field.declarationString.replace(/{[\s\S]*}/gm, "");
        let parameterLink: string = "";
        let paramIndex = fieldString.indexOf(":");
        while (paramIndex < fieldString.indexOf(")")) {
            const wordStartIndex = Math.max(
                fieldString.lastIndexOf("(", paramIndex),
                fieldString.lastIndexOf(" ", paramIndex)) + 1;
            // Remove the variable modifiers for the link.
            parameterLink += "_" + fieldString.substring(wordStartIndex, paramIndex).replace(/\?/gm, "").replace(/\.\.\./gm, "") + "_";
            paramIndex = fieldString.indexOf(":", paramIndex + 1);
        }

        if (parameterLink === "") {
            parameterLink = "__";
        }


        fieldLink = "/" + relativePath + className.toLowerCase() + "#" + field.name + parameterLink;
    } else {
        fieldLink = "/" + relativePath + className.toLowerCase() + "#" + field.name;
    }

    return fieldLink;
}

export function parseDTS(fileName: string, fileContents: string): APISet {
    const node : ts.Node = ts.createSourceFile(
        fileName,
        fileContents,
        ts.ScriptTarget.ES2015,
        true);
    const allClasses: APISet = new APISet();
    parseDTSInternal(node, allClasses);
    return allClasses;
}

function parseDTSInternal(node: ts.Node, allClasses: APISet): void {
    switch (node.kind) {
        case ts.SyntaxKind.InterfaceDeclaration:
            parseDTSTopLevelItem(node as ts.InterfaceDeclaration, allClasses, ClassType.Interface);
            break;
        case ts.SyntaxKind.ClassDeclaration:
            parseDTSTopLevelItem(node as ts.ClassDeclaration, allClasses, ClassType.Class);
            break;
        case ts.SyntaxKind.EnumDeclaration:
            parseDTSTopLevelItem(node as ts.EnumDeclaration, allClasses, ClassType.Enum);
            break;
        case ts.SyntaxKind.PropertySignature:
            parseDTSFieldItem(node as ts.PropertySignature, FieldType.Property);
            break;
        case ts.SyntaxKind.PropertyDeclaration:
            parseDTSFieldItem(node as ts.PropertyDeclaration, FieldType.Property);
            break;
        case ts.SyntaxKind.EnumMember:
            parseDTSFieldItem(node as ts.EnumMember, FieldType.Enum);
            break;
        case ts.SyntaxKind.MethodSignature:
            parseDTSFieldItem(node as ts.MethodSignature, FieldType.Method);
            break;
        case ts.SyntaxKind.MethodDeclaration:
            parseDTSFieldItem(node as ts.MethodDeclaration, FieldType.Method);
            break;
        default:
            // the compiler parses comments after the class/field, therefore this connects to the previous item
            if (node.getText().indexOf("/**") >= 0 &&
                node.getText().indexOf("*/") >= 0 &&
                lastItem !== null &&
                lastItem.comment === "") {
                // clean up spacing as best we can for the diffed d.ts
                lastItem.comment = node.getText().replace(/    \*/g, "*");
                if (lastItem.comment.indexOf("@eventproperty") >= 0) {
                    // events are indistinguishable from properties aside from this tag
                    lastItem.type = FieldType.Event;
                }
            }
    }

    node.getChildren().forEach((element) => {
        parseDTSInternal(element, allClasses);
    });
}

function parseDTSTopLevelItem(
    node: ts.InterfaceDeclaration | ts.ClassDeclaration | ts.EnumDeclaration,
    allClasses: APISet,
    type: ClassType): void {
    topClass = new ClassStruct("export " + type.toLowerCase() + " " + node.name.text, "", type);
    allClasses.addClass(topClass);
    lastItem = topClass;
}

function parseDTSFieldItem(
    node: ts.PropertySignature | ts.PropertyDeclaration | ts.EnumMember | ts.MethodSignature | ts.MethodDeclaration,
    type: FieldType): void {
    // checking for and ignoring mid-method parameters for load()
    if (node.getText().indexOf("expand?") < 0 && node.getText().indexOf("select?") < 0) {
        const newField: FieldStruct = new FieldStruct(node.getText(), "", type, node.name.getText());
        topClass.fields.push(newField);
        lastItem = newField;
    }
}
