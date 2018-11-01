// usage: node WhatsNew "Old d.ts" "New d.ts" "link prefix"
// example: node WhatsNew excel-1_7.d.ts excel-1_8.d.ts javascript/api/excel/excel.

import { readFileSync } from "fs";
import * as fsx from "fs-extra";
import * as ts from "typescript";

enum ClassType {
    Class = "Class",
    Interface = "Interface",
    Enum = "Enum"
}

enum FieldType {
    Property = "Property",
    Method = "Method",
    Event = "Event",
    Enum = "Enum"
}

class FieldStruct {
    declarationString: string;
    comment: string;
    type: FieldType;

    constructor(decString: string, commentString: string, fieldType: FieldType) {
        this.declarationString = decString;
        this.comment = commentString;
        this.type = fieldType;
    }
}

class ClassStruct {
    declarationString: string;
    comment: string;
    fields: FieldStruct[];

    constructor(decString: string, commentString: string) {
        this.declarationString = decString;
        this.comment = commentString;
        this.fields = [];
   }

   copyWithoutFields(): ClassStruct {
       return new ClassStruct(this.declarationString, this.comment);
   }

   sortFields(): void {
       this.fields.sort((a, b) => {
           if (a.declarationString === b.declarationString) {
               return 0;
           } else {
            return a.declarationString < b.declarationString ? -1 : 1;
           }
       });
   }

   getClassName(): string {
       return this.declarationString.substring(this.declarationString.lastIndexOf(" ") + 1);
   }
}

class APISet {
    api: ClassStruct[];
    constructor() {
         this.api = [];
    }

    addClass(clas: ClassStruct): void {
        this.api.push(clas);
    }

    containsClass(clas: ClassStruct): boolean {
        let found: boolean = false;
        this.api.forEach(element => {
            if (element.declarationString === clas.declarationString) {
                found = true;
            }
        });

        return found;
    }

    containsField(field: FieldStruct): boolean {
        let found: boolean = false;
        this.api.forEach(element => {
            element.fields.forEach(thisField => {
                if (thisField.declarationString === field.declarationString) {
                    found = true;
                }
            });
        });

        return found;
    }

    getParentClassShell(field: FieldStruct): ClassStruct {
        let parent: ClassStruct = null;
        this.api.forEach(element => {
            element.fields.forEach(thisField => {
                if (thisField.declarationString === field.declarationString) {
                    parent = element.copyWithoutFields();
                }
            });
        });
        return parent;
    }

    diff(other: APISet): APISet {
        let diffAPI: APISet = new APISet();
        this.api.forEach(element => {
            if (!other.containsClass(element)) {
                diffAPI.addClass(element);
            } else {
                let classShell: ClassStruct = null;
                element.fields.forEach(field => {
                    if (!other.containsField(field)) {
                        if (classShell === null) {
                            classShell = this.getParentClassShell(field);
                            diffAPI.addClass(classShell);
                        }
                        classShell.fields.push(field);
                    }
                });
            }
        });

        return diffAPI;
    }

    getAsDTS(): string {
        this.sort();
        let output: string = "";
        this.api.forEach(clas => {
            output += clas.comment.trim() + "\n";
            output += clas.declarationString + " {\n";
            clas.fields.forEach(field => {
                output += "    " + field.comment + "\n";
                output += "    " + field.declarationString + "\n";
            });
            output += "}\n";
        });
        return output;
    }

    getAsMarkdown(folder: string): string {
        this.sort();
        let output: string = "|Object|What's new|Description|\n";
        output += "|:----|:----|:----|\n";
        this.api.forEach(clas => {
            clas.fields.forEach(field => {
                let className: string = clas.getClassName();
                let tableLine: string = "|[" + className + "](/" + relativePath + className.toLowerCase() + ")|";
                let newItemText: string = field.declarationString.replace(/;/g, "");
                switch (field.type) {
                    case FieldType.Property:
                    newItemText = newItemText.substring(0, newItemText.indexOf(":")).replace("readonly ", "");
                    break;
                    case FieldType.Method:
                    newItemText = newItemText.substring(0, newItemText.indexOf(" =>"));
                    break;
                    case FieldType.Enum:
                    newItemText = newItemText.substring(0, newItemText.indexOf(" ="));
                    break;
                }
                let firstSentenceIndex: number = field.comment.indexOf("* ") + 2;
                tableLine += newItemText + "|";
                tableLine += field.comment.substring(firstSentenceIndex, field.comment.indexOf("\n", firstSentenceIndex)) + "|";
                output += tableLine + "\n";
            });
        });
        return output;
    }

    sort(): void {
        this.api.forEach(element => {
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

let releaseAPI: APISet = new APISet();
let betaAPI: APISet = new APISet();

// read files
const fileNames: string[] = process.argv.slice(2);
let releaseFile: ts.SourceFile = ts.createSourceFile(
    fileNames[0],
    readFileSync(fileNames[0]).toString(),
    ts.ScriptTarget.ES2015,
    true
);
let betaFile: ts.SourceFile = ts.createSourceFile(
    fileNames[1],
    readFileSync(fileNames[1]).toString(),
    ts.ScriptTarget.ES2015,
    true
);

let topClass: ClassStruct = null;
let lastItem: ClassStruct | FieldStruct = null;

parseDTS(releaseFile, releaseAPI);
parseDTS(betaFile, betaAPI);

let diffAPI : APISet = betaAPI.diff(releaseAPI);

let relativePath: string = process.argv[4];
fsx.writeFileSync("WhatsNew.d.ts", diffAPI.getAsDTS());
fsx.writeFileSync("WhatsNew.md", diffAPI.getAsMarkdown(relativePath));

function parseDTS(node: ts.Node, allClasses: APISet): void {
    switch (node.kind) {
        case ts.SyntaxKind.InterfaceDeclaration:
            parseDTSTopLevelItem(<ts.InterfaceDeclaration>node, allClasses, ClassType.Interface);
            break;
        case ts.SyntaxKind.ClassDeclaration:
            parseDTSTopLevelItem(<ts.ClassDeclaration>node, allClasses, ClassType.Class);
            break;
        case ts.SyntaxKind.EnumDeclaration:
            parseDTSTopLevelItem(<ts.EnumDeclaration>node, allClasses, ClassType.Enum);
            break;
        case ts.SyntaxKind.PropertySignature:
            parseDTSFieldItem(<ts.PropertySignature>node, FieldType.Property);
            break;
        case ts.SyntaxKind.PropertyDeclaration:
            parseDTSFieldItem(<ts.PropertyDeclaration>node, FieldType.Property);
            break;
        case ts.SyntaxKind.EnumMember:
            parseDTSFieldItem(<ts.EnumMember>node, FieldType.Enum);
            break;
        case ts.SyntaxKind.MethodSignature:
            parseDTSFieldItem(<ts.MethodSignature>node, FieldType.Method);
            break;
        default:
            if (node.getText().indexOf("/**") >= 0 && node.getText().indexOf("*/") && lastItem !== null && lastItem.comment === "") {
                lastItem.comment = node.getText().replace(/    \*/g, "*");
            }
    }
    node.getChildren().forEach(element => {
        parseDTS(element, allClasses);
    });
}

function parseDTSTopLevelItem(
  node: ts.InterfaceDeclaration | ts.ClassDeclaration | ts.EnumDeclaration,
  allClasses: APISet,
  type: ClassType): void {
    topClass = new ClassStruct("export " + type.toLowerCase() + " " + node.name.text, "");
    allClasses.addClass(topClass);
    lastItem = topClass;
}

function parseDTSFieldItem(
  node: ts.PropertySignature | ts.PropertyDeclaration | ts.EnumMember | ts.MethodSignature,
  type: FieldType): void {
    let newField: FieldStruct = new FieldStruct(node.getText(), "", type);
    topClass.fields.push(newField);
    lastItem = newField;
}

