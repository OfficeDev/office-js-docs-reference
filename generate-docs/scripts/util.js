"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
const fs = require("fs-extra");
require('isomorphic-fetch');
function fetchAndThrowOnError(url, format) {
    return __awaiter(this, void 0, void 0, function* () {
        let response = yield fetch(url);
        if (response.status >= 400) {
            throw new Error(`Bad response from server for URL ${url}`);
        }
        switch (format) {
            case 'text':
                return yield response.text();
            case 'json':
                return yield response.json();
            default:
                throw new Error("Invalid format specified");
        }
    });
}
exports.fetchAndThrowOnError = fetchAndThrowOnError;
function createOrEmptyOutDirectory(path) {
    if (fs.existsSync(path)) {
        fs.emptyDirSync(path);
    }
    else {
        fs.mkdirSync(path);
    }
}
exports.createOrEmptyOutDirectory = createOrEmptyOutDirectory;
function stripSpaces(text) {
    let lines = text.split('\n');
    // Replace each tab with 4 spaces.
    for (let i = 0; i < lines.length; i++) {
        lines[i].replace('\t', '    ');
    }
    let isZeroLengthLine = true;
    let arrayPosition = 0;
    // Remove zero length lines from the beginning of the snippet.
    do {
        let currentLine = lines[arrayPosition];
        if (currentLine.trim() === '') {
            lines.splice(arrayPosition, 1);
        }
        else {
            isZeroLengthLine = false;
        }
    } while (isZeroLengthLine || (arrayPosition === lines.length));
    arrayPosition = lines.length - 1;
    isZeroLengthLine = true;
    // Remove zero length lines from the end of the snippet.
    do {
        let currentLine = lines[arrayPosition];
        if (currentLine.trim() === '') {
            lines.splice(arrayPosition, 1);
            arrayPosition--;
        }
        else {
            isZeroLengthLine = false;
        }
    } while (isZeroLengthLine);
    // Get smallest indent for align left.
    let shortestIndentSize = 1024;
    for (let line of lines) {
        let currentLine = line;
        if (currentLine.trim() !== '') {
            let spaces = line.search(/\S/);
            if (spaces < shortestIndentSize) {
                shortestIndentSize = spaces;
            }
        }
    }
    // Align left
    for (let i = 0; i < lines.length; i++) {
        if (lines[i].length >= shortestIndentSize) {
            lines[i] = lines[i].substring(shortestIndentSize);
        }
    }
    // Convert the array back into a string and return it.
    let finalSetOfLines = '';
    for (let i = 0; i < lines.length; i++) {
        if (i < lines.length - 1) {
            finalSetOfLines += lines[i] + '\n';
        }
        else {
            finalSetOfLines += lines[i];
        }
    }
    return finalSetOfLines;
}
exports.stripSpaces = stripSpaces;
//# sourceMappingURL=util.js.map