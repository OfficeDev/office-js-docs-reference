#!/usr/bin/env node --harmony
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
const util_1 = require("./util");
const fsx = require("fs-extra");
class DtsBuilder {
    constructor() {
        this.slashes = '////////////////////////////////////////////////////////////////';
    }
    extractDtsSection(definitions, beginMarker, endMarker) {
        const definitionsLowercase = definitions.toLowerCase();
        const indexOfBefore = this.indexOfOneAndOnlyOneLine(beginMarker.toLowerCase(), definitionsLowercase, "before");
        const indexOfAfter = this.indexOfOneAndOnlyOneLine(endMarker.toLowerCase(), definitionsLowercase, "after");
        return this.slashes +
            definitions.substring(indexOfBefore, indexOfAfter) +
            this.slashes;
    }
    /** Finds the index of a line containing a particular word -- and ensures that only one such line exists */
    indexOfOneAndOnlyOneLine(needle, haystack, adjustTo) {
        const position = haystack.indexOf(needle);
        if (position < 0) {
            throw new Error(`Could not find "${needle}"`);
        }
        const nextPosition = haystack.indexOf(needle, position + needle.length);
        if (nextPosition > 0) {
            throw new Error(`Expecting one and only one occurence of the word "${needle}"`);
        }
        switch (adjustTo) {
            case "before":
                return haystack.lastIndexOf('\n', position);
            case "after":
                return haystack.indexOf('\n', position) + 1;
            default:
                throw new Error("Invalid position specified");
        }
    }
}
exports.DtsBuilder = DtsBuilder;
(() => __awaiter(this, void 0, void 0, function* () {
    const url = 'https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts';
    let dtsBuilder = new DtsBuilder();
    let definitions = yield util_1.fetchAndThrowOnError(url, "text");
    // fix issues with d.ts file
    definitions = definitions.replace(/^(\s*)(declare namespace)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(declare module)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(namespace)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(class)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(interface)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(module)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(function)(\s+)/gm, `$1export $2$3`)
        .replace(/(\s*)(@param)(\s+)(\w+)(\s)(\s)/g, `$1$2$3$4$5`)
        .replace(/(\s*)(@param)(\s+)(\w+)(\s+)([^\-])/g, `$1$2$3$4$5- $6`);
    // create file: excel.d.ts
    fsx.writeFileSync('../api-extractor-inputs-excel/excel.d.ts', dtsBuilder.extractDtsSection(definitions, "Begin Excel APIs", "End Excel APIs"));
    // create file: office.d.ts
    fsx.writeFileSync('../api-extractor-inputs-office/office.d.ts', dtsBuilder.extractDtsSection(definitions, "Begin Office namespace", "End Office namespace") +
        '\n' +
        '\n' +
        dtsBuilder.extractDtsSection(definitions, "Begin OfficeExtension runtime", "End OfficeExtension runtime"));
    // create file: onenote.d.ts
    fsx.writeFileSync('../api-extractor-inputs-onenote/onenote.d.ts', dtsBuilder.extractDtsSection(definitions, "Begin OneNote APIs", "End OneNote APIs"));
    // create file: outlook.d.ts
    fsx.writeFileSync('../api-extractor-inputs-outlook/outlook.d.ts', dtsBuilder.extractDtsSection(definitions, "Begin Exchange APIs", "End Exchange APIs"));
    // create file: viso.d.ts
    fsx.writeFileSync('../api-extractor-inputs-visio/visio.d.ts', dtsBuilder.extractDtsSection(definitions, "Begin Visio APIs", "End Visio APIs"));
    // create file: word.d.ts
    fsx.writeFileSync('../api-extractor-inputs-word/word.d.ts', dtsBuilder.extractDtsSection(definitions, "Begin Word APIs", "End Word APIs"));
}))();
//# sourceMappingURL=preprocessor.js.map