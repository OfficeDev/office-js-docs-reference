// usage: node WhatsNew "Old d.ts" "New d.ts" "link prefix"
// example: node WhatsNew excel-1_7.d.ts excel-1_8.d.ts javascript/api/excel/excel.

import { readFileSync } from "fs";
import * as fsx from "fs-extra";
import * as ts from "typescript";

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
                return a.declarationString < b.declarationString ? -1 : 1;
            }
        });
    }

    public getClassName(): string {
        return this.declarationString.substring(this.declarationString.lastIndexOf(" ") + 1);
    }
}

class APISet {
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
        let output: string = "";
        this.api.forEach((clas) => {
            // ignore enums
            if (clas.type !== ClassType.Enum) {
                const className = clas.getClassName();
                output += "### [" + className + "](/"
                    + relativePath + className.toLowerCase() + ")\n\n";
                output += extractFirstSentenceFromComment(clas.comment);
                output += "\n\n> [!div class=\"mx-tdCol0BreakAll\"]\n> |Fields|Description|\n> |:---|:---|\n";
                clas.fields.forEach((field) => {
                    // remove unnecessary parts of the declaration string
                    let newItemText = field.declarationString.replace(/;/g, "");
                    newItemText = newItemText.substring(0, newItemText.lastIndexOf(":")).replace("readonly ", "");
                    newItemText = newItemText.replace(/\|/g, "\\|");
                    let tableLine = "> |[" + newItemText + "]("
                        + buildFieldLink(relativePath, className, field) + ")|";

                    tableLine += extractFirstSentenceFromComment(field.comment);
                    output += tableLine + "|\n";
                });
                output += "\n";
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
    let endIndex = commentText.indexOf("\n", firstSentenceIndex);
    if (endIndex === -1) {
        // this is necessary if the comment is a single line (as in collections)
        endIndex = commentText.indexOf("\*/");
    }

    return commentText.substring(firstSentenceIndex, endIndex).trim();
}

function buildFieldLink(relativePath: string, className: string, field: FieldStruct) {
    switch (field.type) {
        case FieldType.Method:
            let parameterLink: string = "";
            let paramIndex = field.declarationString.indexOf(":");
            while (paramIndex > 0 && paramIndex < field.declarationString.indexOf(")")) {
                const wordStartIndex = Math.max(
                    field.declarationString.lastIndexOf("(", paramIndex),
                    field.declarationString.lastIndexOf(" ", paramIndex)) + 1;
                parameterLink += "-" + field.declarationString.substring(wordStartIndex, paramIndex) + "-";
                paramIndex = field.declarationString.indexOf(":", paramIndex + 1);
            }
            return "/" + relativePath + className.toLowerCase() + "#" + field.name + parameterLink;
        default:
            return "/" + relativePath + className.toLowerCase() + "#" + field.name;
    }
}

function fixDTS(definitions: string): string {
    // remove undesirable content, like load, set, data classes, and toJSON
    return definitions
        .replace(/\s*load\(option\?: (Excel|Word|OneNote|Visio)\.Interfaces\.\S*LoadOptions.*\): \S*?;/gm, '')
        .replace(/\*\s*?`load\(option\?: string \| string\[\]\): (Excel|Word|OneNote|Visio)\..*?` - Where option is a comma-delimited string or an array of strings that specify the properties to load\./g, '')
        .replace(/interface .*?LoadOptions \{[^}]*?}/gm, '')
        .replace(/interface .*?Data \{[^}]*?}/gm, '')
        .replace(/load\(option\?\: string \| string\[\]\)\: .*\;/gm, '')
        .replace(/toJSON\(\)\:.*\;/gm, '')
        .replace(/\/\*\* Sets multiple properties.*\s*\*\s*\*.@remarks\s*\*\s*\* This method has the following additional signature:\s*\*\s*\* \`set\(properties:.*\s*\*\s*\* @param.*\s*\*.*\s*\*\/\s*set\(properties:.*\s*\/\*\* Sets multiple properties.*\s*set\(properties:.*;/gm, '')
        .replace(/context\: RequestContext\;/gm, "")
        .replace(/\/\*\* The request context associated with the object\. This connects the add\-in\'s process to the Office host application\'s process\. \*\//gm, "");
}

function parseDTS(node: ts.Node, allClasses: APISet): void {
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
                    // events are indistingushable from properties aside from this tag
                    lastItem.type = FieldType.Event;
                }
            }
    }

    node.getChildren().forEach((element) => {
        parseDTS(element, allClasses);
    });
}

function parseDTSTopLevelItem(
    node: ts.InterfaceDeclaration | ts.ClassDeclaration | ts.EnumDeclaration,
    allClasses: APISet,
    type: ClassType): void {
    //console.log("Creating " + node.name.text);
    topClass = new ClassStruct("export " + type.toLowerCase() + " " + node.name.text, "", type);
    allClasses.addClass(topClass);
    lastItem = topClass;
}

function parseDTSFieldItem(
    node: ts.PropertySignature | ts.PropertyDeclaration | ts.EnumMember | ts.MethodSignature | ts.MethodDeclaration,
    type: FieldType): void {
    if (node.getText().indexOf("expand?") < 0 && node.getText().indexOf("select?") < 0) {
        // checking for and ignoring mid-method parameters for load()
        const newField: FieldStruct = new FieldStruct(node.getText(), "", type, node.name.getText());
        //console.log("Adding " + newField.name + " to " + topClass.getClassName());
        topClass.fields.push(newField);
        lastItem = newField;
    }
}

// capturing these because of eccentrities with the compiler ordering
let topClass: ClassStruct = null;
let lastItem: ClassStruct | FieldStruct = null;

(() => {

    const releaseAPI: APISet = new APISet();
    const betaAPI: APISet = new APISet();

    // read files
    const fileNames: string[] = process.argv.slice(2);
    const releaseFile: ts.SourceFile = ts.createSourceFile(
        "Release",
        fixDTS(readFileSync(fileNames[0]).toString()),
        ts.ScriptTarget.ES2015,
        true);
    const betaFile: ts.SourceFile = ts.createSourceFile(
        "Preview",
        fixDTS(readFileSync(fileNames[1]).toString()),
        ts.ScriptTarget.ES2015,
        true);
    parseDTS(releaseFile, releaseAPI);
    parseDTS(betaFile, betaAPI);

    fsx.writeFileSync("Release.d.ts", releaseAPI.getAsDTS());
    fsx.writeFileSync("Beta.d.ts", betaAPI.getAsDTS());


    const diffAPI: APISet = betaAPI.diff(releaseAPI);

    const relativePath: string = process.argv[4];
    fsx.writeFileSync("WhatsNew.d.ts", diffAPI.getAsDTS());
    fsx.writeFileSync("WhatsNew.md", diffAPI.getAsMarkdown(relativePath));
})();
