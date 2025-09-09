"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const ts = require("typescript");
// capturing these because of eccentricities with the compiler ordering
let topClass = null;
let lastItem = null;
var ClassType;
(function (ClassType) {
    ClassType["Class"] = "Class";
    ClassType["Interface"] = "Interface";
    ClassType["Enum"] = "Enum";
})(ClassType || (ClassType = {}));
var FieldType;
(function (FieldType) {
    FieldType["Property"] = "Property";
    FieldType["Method"] = "Method";
    FieldType["Event"] = "Event";
    FieldType["Enum"] = "Enum";
})(FieldType || (FieldType = {}));
class FieldStruct {
    constructor(decString, commentString, fieldType, fieldName) {
        this.declarationString = decString;
        this.comment = commentString;
        this.type = fieldType;
        this.name = fieldName;
    }
}
class ClassStruct {
    constructor(decString, commentString, classType) {
        this.declarationString = decString;
        this.comment = commentString;
        this.type = classType;
        this.fields = [];
    }
    copyWithoutFields() {
        return new ClassStruct(this.declarationString, this.comment, this.type);
    }
    sortFields() {
        this.fields.sort((a, b) => {
            if (a.declarationString === b.declarationString) {
                return 0;
            }
            else {
                return a.declarationString.replace("readonly ", "") < b.declarationString.replace("readonly ", "") ? -1 : 1;
            }
        });
    }
    getClassName() {
        return this.declarationString.substring(this.declarationString.lastIndexOf(" ") + 1);
    }
}
class APISet {
    constructor() {
        this.api = [];
    }
    addClass(clas) {
        this.api.push(clas);
    }
    containsClass(clas) {
        let found = false;
        this.api.forEach((element) => {
            if (element.declarationString === clas.declarationString) {
                found = true;
            }
        });
        return found;
    }
    containsField(clas, field) {
        let found = false;
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
    diff(other) {
        const diffAPI = new APISet();
        this.api.forEach((element) => {
            if (other.containsClass(element)) {
                let classShell = null;
                element.fields.forEach((field) => {
                    if (!other.containsField(element, field)) {
                        if (classShell === null) {
                            classShell = element.copyWithoutFields();
                            diffAPI.addClass(classShell);
                        }
                        classShell.fields.push(field);
                    }
                });
            }
            else {
                diffAPI.addClass(element);
            }
        });
        return diffAPI;
    }
    getAsDTS() {
        this.sort();
        const output = [];
        this.api.forEach((clas) => {
            output.push(clas.comment.trim());
            output.push(clas.declarationString + " {");
            clas.fields.forEach((field) => {
                output.push("    " + field.comment);
                if (field.type === FieldType.Enum) {
                    output.push("    " + field.declarationString + ",");
                }
                else {
                    output.push("    " + field.declarationString);
                }
            });
            output.push("}");
        });
        return output.join("\n");
    }
    getAsMarkdown(relativePath) {
        this.sort();
        // table header
        let output = "| Class | Fields | Description |\n|:---|:---|:---|\n";
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
                // Ignore the following:
                // - String literal overloads.
                // - `load`, `set`, `track`, `untrack`, and `toJSON` methods
                // - The `context` property.
                // - Static fields.
                let filteredFields = clas.fields.filter((field) => {
                    let isLiteral = field.declarationString.search(/([a-zA-Z]+)(\??:)([\n]?([ |]*\"[\w]*\"[|,\n]*)+?)([ ]*[\),])/g) >= 0;
                    return (!isLiteral &&
                        field.name !== "load" &&
                        field.name !== "set" &&
                        field.name !== "toJSON" &&
                        field.name !== "context" &&
                        field.name !== "track" &&
                        field.name !== "untrack" &&
                        !field.declarationString.includes("static "));
                });
                let first = true;
                if (filteredFields.length > 0) {
                    filteredFields.forEach((field) => {
                        if (first) {
                            first = false;
                        }
                        else {
                            output += "||";
                        }
                        // remove unnecessary parts of the declaration string
                        let newItemText = field.declarationString.replace(/;/g, "");
                        if (field.type === FieldType.Property) {
                            // Remove the optional modifier and type.
                            newItemText = newItemText.replace(/\?/g, "");
                            newItemText = newItemText.substring(0, newItemText.indexOf(":"));
                        }
                        else {
                            // Remove the return type.
                            newItemText = newItemText.substring(0, newItemText.lastIndexOf(":"));
                        }
                        newItemText = newItemText.replace("readonly ", "");
                        newItemText = newItemText.replace(/\|/g, "\\|").replace(/\n|\t/gm, "");
                        newItemText = newItemText.replace(/[\s][\s]+/g, " ").replace(/\( /g, "(").replace(/ \)/g, ")").replace(/,\)/g, ")").replace(/([\w]\??: )\\\| /g, "$1"); // dprint formatting quirks
                        newItemText = newItemText.replace(/\<any\>/g, "");
                        let tableLine = "[" + newItemText + "]("
                            + buildFieldLink(relativePath, className, field) + ")|";
                        tableLine += removeAtLink(extractFirstSentenceFromComment(field.comment));
                        output += tableLine + "|\n";
                    });
                }
                else {
                    output += "||\n";
                }
            }
        });
        return output;
    }
    sort() {
        this.api.forEach((element) => {
            element.sortFields();
        });
        this.api.sort((a, b) => {
            if (a.getClassName() === b.getClassName()) {
                return 0;
            }
            else {
                return a.getClassName() < b.getClassName() ? -1 : 1;
            }
        });
    }
}
exports.APISet = APISet;
function extractFirstSentenceFromComment(commentText) {
    const firstSentenceIndex = commentText.indexOf("* ") + 2;
    const multiSentenceEndIndex = commentText.indexOf(". ", firstSentenceIndex);
    const lineBreakEndIndex = commentText.indexOf("\n", firstSentenceIndex);
    const singleLineEndIndex = commentText.indexOf("\*/", firstSentenceIndex);
    let endIndex;
    if (multiSentenceEndIndex > 0 && lineBreakEndIndex > 0) {
        endIndex = Math.min(multiSentenceEndIndex + 1, lineBreakEndIndex);
    }
    else if (multiSentenceEndIndex === -1 && lineBreakEndIndex === -1) {
        endIndex = singleLineEndIndex;
    }
    else {
        endIndex = Math.max(multiSentenceEndIndex + 1, lineBreakEndIndex);
    }
    return commentText.substring(firstSentenceIndex, endIndex).trim();
}
function removeAtLink(commentText) {
    // Replace links with the format "{@link Foo}" with "Foo".
    commentText = commentText.replace(/{@link ([^|]*?)}/gm, "$1");
    // Replace links with the format "{@link Foo | URL}" with "[Foo](URL)".
    commentText = commentText.replace(/{@link ([^}]*?) \| (http.*?)}/gm, "[$1]($2)");
    return commentText;
}
function buildFieldLink(relativePath, className, field) {
    // Build the standard link anchor format based on host.
    let anchorPrefix = relativePath.substring(relativePath.lastIndexOf("/") + 1, relativePath.lastIndexOf("."));
    anchorPrefix = (relativePath.indexOf("outlook") > 0 ? "outlook" : anchorPrefix) + "-" + anchorPrefix + "-";
    let fieldLink = "/" + relativePath.replace("api/outlook/outlook", "api/outlook/office") + className.toLowerCase() + "#" + anchorPrefix + className.toLowerCase() + "-" + field.name.toLowerCase() + (field.type === FieldType.Method ? "-member(1)" : "-member");
    return fieldLink;
}
function parseDTS(fileName, fileContents) {
    const node = ts.createSourceFile(fileName, fileContents, ts.ScriptTarget.ES2015, true);
    const allClasses = new APISet();
    parseDTSInternal(node, allClasses);
    return allClasses;
}
exports.parseDTS = parseDTS;
function parseDTSInternal(node, allClasses) {
    switch (node.kind) {
        case ts.SyntaxKind.InterfaceDeclaration:
            parseDTSTopLevelItem(node, allClasses, ClassType.Interface);
            break;
        case ts.SyntaxKind.ClassDeclaration:
            parseDTSTopLevelItem(node, allClasses, ClassType.Class);
            break;
        case ts.SyntaxKind.EnumDeclaration:
            parseDTSTopLevelItem(node, allClasses, ClassType.Enum);
            break;
        case ts.SyntaxKind.PropertySignature:
            parseDTSFieldItem(node, FieldType.Property);
            break;
        case ts.SyntaxKind.PropertyDeclaration:
            parseDTSFieldItem(node, FieldType.Property);
            break;
        case ts.SyntaxKind.EnumMember:
            parseDTSFieldItem(node, FieldType.Enum);
            break;
        case ts.SyntaxKind.MethodSignature:
            parseDTSFieldItem(node, FieldType.Method);
            break;
        case ts.SyntaxKind.MethodDeclaration:
            parseDTSFieldItem(node, FieldType.Method);
            break;
        case ts.SyntaxKind.TypeLiteral:
            return;
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
function parseDTSTopLevelItem(node, allClasses, type) {
    topClass = new ClassStruct("export " + type.toLowerCase() + " " + node.name.text, "", type);
    allClasses.addClass(topClass);
    lastItem = topClass;
}
function parseDTSFieldItem(node, type) {
    const newField = new FieldStruct(node.getText(), "", type, node.name.getText());
    topClass.fields.push(newField);
    lastItem = newField;
}
//# sourceMappingURL=dts-utilities.js.map