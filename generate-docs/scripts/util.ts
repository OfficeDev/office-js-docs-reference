import * as fs from "fs-extra";
import * as ts from "typescript";
require('isomorphic-fetch');

export async function fetchAndThrowOnError(url: string, format: 'text'): Promise<string>;
export async function fetchAndThrowOnError<T>(url: string, format: 'json'): Promise<T>;
export async function fetchAndThrowOnError(url: string, format: 'text' | 'json') {
    let response = await fetch(url);
    if (response.status >= 400) {
        throw new Error(`Bad response from server for URL ${url}`);
    }

    switch (format) {
        case 'text':
            return await response.text();
        case 'json':
            return await response.json();
        default:
            throw new Error("Invalid format specified");
    }
}

export class DtsBuilder {
    private slashes = '////////////////////////////////////////////////////////////////';

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
            throw new Error(`Expecting one and only one occurrence of the word "${needle}"`);
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

export function createOrEmptyOutDirectory(path: string) {
    if (fs.existsSync(path)) {
        fs.emptyDirSync(path);
    } else {
        fs.mkdirSync(path);
    }
}

export function stripSpaces(text: string) {
    let lines: string[] = text.split('\n');

    // Replace each tab with 4 spaces.
    for (let i: number = 0; i < lines.length; i++) {
        lines[i].replace('\t', '    ');
    }

    let isZeroLengthLine: boolean = true;
    let arrayPosition: number = 0;

    // Remove zero length lines from the beginning of the snippet.
    do {
        let currentLine: string = lines[arrayPosition];
        if (currentLine.trim() === '') {
            lines.splice(arrayPosition, 1);
        } else {
            isZeroLengthLine = false;
        }
    } while (isZeroLengthLine || (arrayPosition === lines.length));

    arrayPosition = lines.length - 1;
    isZeroLengthLine = true;

    // Remove zero length lines from the end of the snippet.
    do {
        let currentLine: string = lines[arrayPosition];
        if (currentLine.trim() === '') {
            lines.splice(arrayPosition, 1);
            arrayPosition--;
        } else {
            isZeroLengthLine = false;
        }
    } while (isZeroLengthLine);

    // Get smallest indent for align left.
    let shortestIndentSize: number = 1024;
    for (let line of lines) {
        let currentLine: string = line;
        if (currentLine.trim() !== '') {
            let spaces: number = line.search(/\S/);
            if (spaces < shortestIndentSize) {
                shortestIndentSize = spaces;
            }
        }
    }

    // Align left
    for (let i: number = 0; i < lines.length; i++) {
        if (lines[i].length >= shortestIndentSize) {
            lines[i] = lines[i].substring(shortestIndentSize);
        }
    }

    // Convert the array back into a string and return it.
    let finalSetOfLines: string = '';
    for (let i: number = 0; i < lines.length; i++) {
        if (i < lines.length - 1) {
            finalSetOfLines += lines[i] + '\n';
        }
        else {
            finalSetOfLines += lines[i];
        }
    }
    return finalSetOfLines;
}

export function generateEnumList(dtsFile: string) : string[] {
    const releaseFile: ts.SourceFile = ts.createSourceFile(
        "office",
        dtsFile,
        ts.ScriptTarget.ES2015,
        true);
    let enumList = [];
    lookForEnums(releaseFile, enumList);
    return enumList;
}

function lookForEnums(node: ts.Node, enumList: string[]): void {
    switch (node.kind) {
        case ts.SyntaxKind.EnumDeclaration:
            const enumNode = node as ts.EnumDeclaration;
            const enumName = enumNode.name.getText();
            enumList.push(enumName);
            break;
    }

    node.getChildren().forEach((element) => {
        lookForEnums(element, enumList);
    });
}

