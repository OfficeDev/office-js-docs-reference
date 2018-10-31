"use strict";
// usage: node WhatsNew "Old d.ts" "New d.ts" "link prefix"
// example: node WhatsNew excel-1_7.d.ts excel-1_8.d.ts javascript/api/excel/excel.
exports.__esModule = true;
var fs_1 = require("fs");
var fsx = require("fs-extra");
var ts = require("typescript");
var FieldType;
(function (FieldType) {
    FieldType["Property"] = "Property";
    FieldType["Method"] = "Method";
    FieldType["Event"] = "Event";
    FieldType["Enum"] = "Enum";
})(FieldType || (FieldType = {}));
var FieldStruct = /** @class */ (function () {
    function FieldStruct(decString, commentString, fieldType) {
        this.declarationString = decString;
        this.comment = commentString;
        this.type = fieldType;
    }
    return FieldStruct;
}());
var ClassStruct = /** @class */ (function () {
    function ClassStruct(decString, commentString) {
        this.declarationString = decString;
        this.comment = commentString;
        this.fields = [];
    }
    ClassStruct.prototype.copyWithoutFields = function () {
        return new ClassStruct(this.declarationString, this.comment);
    };
    ClassStruct.prototype.sortFields = function () {
        this.fields.sort(function (a, b) {
            if (a.declarationString === b.declarationString) {
                return 0;
            }
            else {
                return a.declarationString < b.declarationString ? -1 : 1;
            }
        });
    };
    ClassStruct.prototype.getClassName = function () {
        return this.declarationString.substring(this.declarationString.lastIndexOf(" ") + 1);
    };
    return ClassStruct;
}());
var APISet = /** @class */ (function () {
    function APISet() {
        this.api = [];
    }
    APISet.prototype.addClass = function (clas) {
        this.api.push(clas);
    };
    APISet.prototype.containsClass = function (clas) {
        var found = false;
        this.api.forEach(function (element) {
            if (element.declarationString === clas.declarationString) {
                found = true;
            }
        });
        return found;
    };
    APISet.prototype.containsField = function (field) {
        var found = false;
        this.api.forEach(function (element) {
            element.fields.forEach(function (thisField) {
                if (thisField.declarationString === field.declarationString) {
                    found = true;
                }
            });
        });
        return found;
    };
    APISet.prototype.getParentClassShell = function (field) {
        var parent = null;
        this.api.forEach(function (element) {
            element.fields.forEach(function (thisField) {
                if (thisField.declarationString === field.declarationString) {
                    parent = element.copyWithoutFields();
                }
            });
        });
        return parent;
    };
    APISet.prototype.diff = function (other) {
        var _this = this;
        var diffAPI = new APISet();
        this.api.forEach(function (element) {
            if (!other.containsClass(element)) {
                diffAPI.addClass(element);
            }
            else {
                var classShell_1 = null;
                element.fields.forEach(function (field) {
                    if (!other.containsField(field)) {
                        if (classShell_1 === null) {
                            classShell_1 = _this.getParentClassShell(field);
                            diffAPI.addClass(classShell_1);
                        }
                        classShell_1.fields.push(field);
                    }
                });
            }
        });
        return diffAPI;
    };
    APISet.prototype.getAsDTS = function () {
        this.sort();
        var output = "";
        this.api.forEach(function (clas) {
            output += clas.comment.trim() + "\n";
            output += clas.declarationString + " {\n";
            clas.fields.forEach(function (field) {
                output += "    " + field.comment + "\n";
                output += "    " + field.declarationString + "\n";
            });
            output += "}\n";
        });
        return output;
    };
    APISet.prototype.getAsMarkdown = function (folder) {
        this.sort();
        var output = "|Object|What's new|Description|\n";
        output += "|:----|:----|:----|\n";
        this.api.forEach(function (clas) {
            clas.fields.forEach(function (field) {
                var className = clas.getClassName();
                var tableLine = "|[" + className + "](/" + relativePath + className.toLowerCase() + ")|";
                var newItemText = field.declarationString.replace(/;/g, "");
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
                var firstSentenceIndex = field.comment.indexOf("* ") + 2;
                tableLine += newItemText + "|";
                tableLine += field.comment.substring(firstSentenceIndex, field.comment.indexOf("\n", firstSentenceIndex)) + "|";
                output += tableLine + "\n";
            });
        });
        return output;
    };
    APISet.prototype.sort = function () {
        this.api.forEach(function (element) {
            element.sortFields();
        });
        this.api.sort(function (a, b) {
            if (a.getClassName() === b.getClassName()) {
                return 0;
            }
            else {
                return a.getClassName() < b.getClassName() ? -1 : 1;
            }
        });
    };
    return APISet;
}());
var releaseAPI = new APISet();
var betaAPI = new APISet();
// read files
var fileNames = process.argv.slice(2);
var releaseFile = ts.createSourceFile(fileNames[0], fs_1.readFileSync(fileNames[0]).toString(), ts.ScriptTarget.ES2015, true);
var betaFile = ts.createSourceFile(fileNames[1], fs_1.readFileSync(fileNames[1]).toString(), ts.ScriptTarget.ES2015, true);
var lastItem = null;
parseDTS(releaseFile, releaseAPI, null);
parseDTS(betaFile, betaAPI, null);
var diffAPI = betaAPI.diff(releaseAPI);
var relativePath = process.argv[4];
fsx.writeFileSync("WhatsNew.d.ts", diffAPI.getAsDTS());
fsx.writeFileSync("WhatsNew.md", diffAPI.getAsMarkdown(relativePath));
function parseDTS(node, allClasses, topClass) {
    switch (node.kind) {
        case ts.SyntaxKind.InterfaceDeclaration:
            var interfaceDeclaration = node;
            topClass = new ClassStruct("export interface " + interfaceDeclaration.name.text, "");
            allClasses.addClass(topClass);
            lastItem = topClass;
            break;
        case ts.SyntaxKind.ClassDeclaration:
            var classDeclaration = node;
            topClass = new ClassStruct("export class " + classDeclaration.name.text, "");
            allClasses.addClass(topClass);
            lastItem = topClass;
            break;
        case ts.SyntaxKind.EnumDeclaration:
            var enumDeclaration = node;
            topClass = new ClassStruct("export enum " + enumDeclaration.name.text, "");
            allClasses.addClass(topClass);
            lastItem = topClass;
            break;
        case ts.SyntaxKind.PropertyDeclaration:
            var propSignature = node;
            var newProp = new FieldStruct(propSignature.getText(), "", FieldType.Property);
            topClass.fields.push(newProp);
            lastItem = newProp;
            break;
        case ts.SyntaxKind.EnumMember:
            var enumSignature = node;
            var newEnum = new FieldStruct(enumSignature.getText(), "", FieldType.Enum);
            topClass.fields.push(newEnum);
            lastItem = newEnum;
            break;
        case ts.SyntaxKind.MethodSignature:
            var methodSignature = node;
            var newMethod = new FieldStruct(methodSignature.getText(), "", FieldType.Method);
            topClass.fields.push(newMethod);
            lastItem = newMethod;
            break;
        default:
            if (node.getText().indexOf("/**") >= 0 && node.getText().indexOf("*/") && lastItem !== null && lastItem.comment === "") {
                lastItem.comment = node.getText().replace(/    \*/g, "*");
            }
    }
    node.getChildren().forEach(function (element) {
        parseDTS(element, allClasses, topClass);
    });
}
