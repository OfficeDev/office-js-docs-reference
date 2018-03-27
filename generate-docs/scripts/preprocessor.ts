#!/usr/bin/env node --harmony

import { fetchAndThrowOnError } from './util';
import * as fsx from 'fs-extra';

export class DtsBuilder {
    private slashes = '////////////////////////////////////////////////////////////////';

    // public replaceDtsSection(definitions: string, beginMarker: string, endMarker: string, content: string): string {
    //     const definitionsLowercase = definitions.toLowerCase();

    //     const indexOfBefore = this.indexOfOneAndOnlyOneLine(
    //         beginMarker.toLowerCase(), definitionsLowercase, "before");
    //     const indexOfAfter = this.indexOfOneAndOnlyOneLine(
    //         endMarker.toLowerCase(), definitionsLowercase, "after");

    //     const before = definitions.substring(0, indexOfBefore);
    //     const after = definitions.substring(indexOfAfter);

    //     return before +
    //         '\n' + this.makeHeader(beginMarker) +
    //         '\n' + this.slashes +
    //         '\n' +
    //         '\n' +
    //         content +
    //         '\n' +
    //         '\n' + this.slashes +
    //         '\n' + this.makeHeader(endMarker) +
    //         '\n' + after;
    // }

    public extractDtsSection(definitions: string, beginMarker: string, endMarker: string): string {
        const definitionsLowercase = definitions.toLowerCase();

        const indexOfBefore = this.indexOfOneAndOnlyOneLine(
            beginMarker.toLowerCase(), definitionsLowercase, "before");
        const indexOfAfter = this.indexOfOneAndOnlyOneLine(
            endMarker.toLowerCase(), definitionsLowercase, "after");

        return this.slashes +
            definitions.substring(indexOfBefore, indexOfAfter) + 
            this.slashes;
    }

    /** Finds the index of a line containing a particular word -- and ensures that only one such line exists */
    private indexOfOneAndOnlyOneLine(needle: string, haystack: string, adjustTo: "before" | "after"): number {
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

    // private makeHeader(text: string): string {
    //     text = ' ' + text + ' ';
    //     const textWithPrefix = this.slashes.substr(0, Math.floor((this.slashes.length - text.length) / 2)) + text;
    //     return textWithPrefix + this.slashes.substr(0, this.slashes.length - textWithPrefix.length);
    // }
}

(async () => {
    const url = 'https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts';

    let dtsBuilder = new DtsBuilder();
    let definitions = await fetchAndThrowOnError(url, "text");

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

    // // remove the OfficeCore section from the d.ts file
    // definitions = dtsBuilder.replaceDtsSection(
    //     definitions,
    //     "Begin OfficeCore",
    //     "End OfficeCore",
    //     ""
    // );

    // create file: excel.d.ts
    fsx.writeFileSync(
        '../api-extractor-inputs-excel/excel.d.ts', 
        dtsBuilder.extractDtsSection(definitions, "Begin Excel APIs", "End Excel APIs")
    );

    // create file: office.d.ts
    fsx.writeFileSync(
        '../api-extractor-inputs-office/office.d.ts', 
        dtsBuilder.extractDtsSection(definitions, "Begin Office namespace", "End Office namespace") + 
        '\n' + 
        '\n' +
        dtsBuilder.extractDtsSection(definitions, "Begin OfficeExtension runtime", "End OfficeExtension runtime")
    );

    // create file: onenote.d.ts
    fsx.writeFileSync(
        '../api-extractor-inputs-onenote/onenote.d.ts', 
        dtsBuilder.extractDtsSection(definitions, "Begin OneNote APIs", "End OneNote APIs")
    );

    // create file: outlook.d.ts
    fsx.writeFileSync(
        '../api-extractor-inputs-outlook/outlook.d.ts', 
        dtsBuilder.extractDtsSection(definitions, "Begin Exchange APIs", "End Exchange APIs")
    );

    // create file: viso.d.ts
    fsx.writeFileSync(
        '../api-extractor-inputs-visio/visio.d.ts', 
        dtsBuilder.extractDtsSection(definitions, "Begin Visio APIs", "End Visio APIs")
    );

    // create file: word.d.ts
    fsx.writeFileSync(
        '../api-extractor-inputs-word/word.d.ts', 
        dtsBuilder.extractDtsSection(definitions, "Begin Word APIs", "End Word APIs")
    );

})();



