#!/usr/bin/env node --harmony

import { fetchAndThrowOnError } from './util';
// import * as fsx from 'fs-extra';

(async () => {
    const url = 'https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts';
    //const slashes = '////////////////////////////////////////////////////////////////';

    // let dtsBuilder = new DtsBuilder();
    let fileContent = await fetchAndThrowOnError(url, "text");

    // add 'export' keyword and hyphen to @param descriptions where necessary
    fileContent = fileContent.replace(/^(\s*)(declare namespace)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(declare module)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(namespace)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(class)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(interface)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(module)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(function)(\s+)/gm, `$1export $2$3`)
        .replace(/(\s*)(@param)(\s+)(\w+)(\s)(\s)/g, `$1$2$3$4$5`)
        .replace(/(\s*)(@param)(\s+)(\w+)(\s+)([^\-])/g, `$1$2$3$4$5- $6`);

    // write file

    // fileContent = dtsBuilder.replaceDtsSection(
    //     fileContent,
    //     "Begin OfficeExtension runtime",
    //     "End OfficeExtension runtime",
    //     [
    //         fs.readFileSync(`${folder}\\IntelliSense_Partial\\officeextension.runtime.manual.d.ts`).toString(),
    //         fs.readFileSync(`${folder}\\IntelliSense_Partial\\office.core.d.ts`).toString()
    //     ].join("\n\n\n")
    // );
})();



export class DtsBuilder {
    private slashes = '////////////////////////////////////////////////////////////////';

    public replaceDtsSection(definitions: string, beginMarker: string, endMarker: string, content: string): string {
        const definitionsLowercase = definitions.toLowerCase();

        const indexOfBefore = this.indexOfOneAndOnlyOneLine(
            beginMarker.toLowerCase(), definitionsLowercase, "before");
        const indexOfAfter = this.indexOfOneAndOnlyOneLine(
            endMarker.toLowerCase(), definitionsLowercase, "after");

        const before = definitions.substring(0, indexOfBefore);
        const after = definitions.substring(indexOfAfter);

        return before +
            '\n' + this.makeHeader(beginMarker) +
            '\n' + this.slashes +
            '\n' +
            '\n' +
            content +
            '\n' +
            '\n' + this.slashes +
            '\n' + this.makeHeader(endMarker) +
            '\n' + after;
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

    private makeHeader(text: string): string {
        text = ' ' + text + ' ';
        const textWithPrefix = this.slashes.substr(0, Math.floor((this.slashes.length - text.length) / 2)) + text;
        return textWithPrefix + this.slashes.substr(0, this.slashes.length - textWithPrefix.length);
    }
}
