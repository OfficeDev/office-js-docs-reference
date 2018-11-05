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

    constructor(decString: string, commentString: string, fieldType: FieldType) {
        this.declarationString = decString;
        this.comment = commentString;
        this.type = fieldType;
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

    public containsField(field: FieldStruct): boolean {
        let found: boolean = false;
        this.api.forEach((element) => {
            element.fields.forEach((thisField) => {
                if (thisField.declarationString === field.declarationString) {
                    found = true;
                }
            });
        });

        return found;
    }

    // used to construct a partial class for the diff
    public getParentClassShell(field: FieldStruct): ClassStruct {
        let parent: ClassStruct = null;
        this.api.forEach((element) => {
            element.fields.forEach((thisField) => {
                if (thisField.declarationString === field.declarationString) {
                    parent = element.copyWithoutFields();
                }
            });
        });
        return parent;
    }

    // finds the new fields and classes
    public diff(other: APISet): APISet {
        const diffAPI: APISet = new APISet();
        this.api.forEach((element) => {
            if (other.containsClass(element)) {
                let classShell: ClassStruct = null;
                element.fields.forEach((field) => {
                    if (!other.containsField(field)) {
                        if (classShell === null) {
                            classShell = this.getParentClassShell(field);
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
                output.push("    " + field.declarationString);
            });
            output.push("}");
        });
        return output.join("\n");
    }

    public getAsMarkdown(relativePath: string): string {
        this.sort();
        // table header
        let output: string = "|Object|What's new|Description|\n";
        output += "|:----|:----|:----|\n";
        this.api.forEach((clas) => {
            // ignore enums
            if (clas.type !== ClassType.Enum) {
                clas.fields.forEach((field) => {
                    const className: string = clas.getClassName();
                    let tableLine: string = "|[" + className + "](/" + relativePath + className.toLowerCase() + ")|";
                    // remove unnecessary parts of the declaration string
                    let newItemText: string = field.declarationString.replace(/;/g, "");
                    switch (field.type) {
                        case FieldType.Property:
                            newItemText = newItemText.substring(0, newItemText.indexOf(":")).replace("readonly ", "")
                             + " (property)";
                            break;
                        case FieldType.Event:
                            newItemText = newItemText.substring(0, newItemText.indexOf(":")).replace("readonly ", "")
                            + " (event)";
                            break;
                    }

                    const firstSentenceIndex: number = field.comment.indexOf("* ") + 2;
                    let endIndex: number = field.comment.indexOf("\n", firstSentenceIndex);
                    if (endIndex === -1) {
                        // this is necessary if the comment is a single line (as in collections)
                        endIndex = field.comment.indexOf("\*/");
                    }

                    tableLine += newItemText + "|";
                    tableLine += field.comment.substring(firstSentenceIndex, endIndex) + "|";
                    output += tableLine + "\n";
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
        default:
            // the compiler parses comments after the class/field, thereof this connects to the previous item
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
    topClass = new ClassStruct("export " + type.toLowerCase() + " " + node.name.text, "", type);
    allClasses.addClass(topClass);
    lastItem = topClass;
}

function parseDTSFieldItem(
    node: ts.PropertySignature | ts.PropertyDeclaration | ts.EnumMember | ts.MethodSignature,
    type: FieldType): void {
    // checking for and ignoring mid-method parameters for load()
    if (node.getText().indexOf("expand?") < 0 && node.getText().indexOf("select?") < 0) {
        const newField: FieldStruct = new FieldStruct(node.getText(), "", type);
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
        fileNames[0],
        readFileSync(fileNames[0]).toString(),
        ts.ScriptTarget.ES2015,
        true);
    const betaFile: ts.SourceFile = ts.createSourceFile(
        fileNames[1],
        readFileSync(fileNames[1]).toString(),
        ts.ScriptTarget.ES2015,
        true);

    parseDTS(releaseFile, releaseAPI);
    parseDTS(betaFile, betaAPI);

    const diffAPI: APISet = betaAPI.diff(releaseAPI);

    const relativePath: string = process.argv[4];
    fsx.writeFileSync("WhatsNew.d.ts", diffAPI.getAsDTS());
    fsx.writeFileSync("WhatsNew.md", diffAPI.getAsMarkdown(relativePath));
})();
