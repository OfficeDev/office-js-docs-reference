import * as fsx from "fs-extra";

if (process.argv.length !== 5 || process.argv.find((x: string) => {return x === "-?"})) {
    console.log("usage: node version-remover [source d.ts] [API set name] [output file name]");
    console.log("example: node version-remover excel.d.ts \"ExcelApi 1.8\" excel_1_7.d.ts");
    process.exit(0);
}

console.log("Version Remover - Creating " + process.argv[4]);
let wholeDTS = fsx.readFileSync(process.argv[2]).toString();
let declarationString;
// find the API tag
let indexOfApiSetTag = wholeDTS.indexOf("Api set: " + process.argv[3]);
while (indexOfApiSetTag >= 0) {
    // find the comment block around the API tag
    let commentStart = wholeDTS.lastIndexOf("/**", indexOfApiSetTag);
    let commentEnd = wholeDTS.indexOf("*/", indexOfApiSetTag);
    commentEnd =  wholeDTS.indexOf("\n", commentEnd) + 1; // Account for newline and ending characters.

    // the declaration string is the line following the comment
    declarationString = wholeDTS.substring(commentEnd, wholeDTS.indexOf("\n", commentEnd));
    let endPosition = commentEnd + declarationString.length;
    if (declarationString.indexOf("class") >= 0 || declarationString.indexOf("enum") >= 0 || declarationString.indexOf("interface") >= 0) {
        endPosition = Math.max(wholeDTS.indexOf("}\r\n", commentEnd), wholeDTS.indexOf("}\n", commentEnd));
    } else {
        endPosition = getDeclarationEnd(wholeDTS, commentEnd);
    }

    if (endPosition === -1) {
        endPosition = commentEnd;
    }
    wholeDTS = wholeDTS.substring(0, commentStart) + wholeDTS.substring(endPosition + 1);
    indexOfApiSetTag = wholeDTS.indexOf("Api set: " + process.argv[3]);
}

/* Add necessary custom logic here*/

if (process.argv[3] === "ExcelApi 1.11") {
    console.log("Address CommentRichContent reference for when removing ExcelApi 1.11");
    wholeDTS = wholeDTS.replace(/add\(content: CommentRichContent \| string,/g, "add(content: string,").
                replace(/add\(cellAddress: Range \| string, content: CommentRichContent \| string,/g, "add(cellAddress: Range | string, content: string,");
}

fsx.writeFileSync(process.argv[4], wholeDTS);


function getDeclarationEnd(wholeDts: string, startIndex: number): number {
    let nextSemicolon = wholeDTS.indexOf(";", startIndex);
    let nextNewLine = wholeDTS.indexOf("\n", startIndex);
    let nextStartBrace = wholeDTS.indexOf("{", startIndex);
    let nextEndBrace = wholeDTS.indexOf("}", startIndex);
    
    // Figure out if the declaration has an internal class.
    if (nextSemicolon < nextNewLine) {
        return nextSemicolon;
    } else if (nextStartBrace > nextNewLine) {
        // No semicolon or braces means this is an enum member.
        return nextNewLine;
    } else {
        // The declaration is on multiple lines, likely due to an internal class.
        wholeDTS.substring(startIndex, wholeDTS.indexOf(";", nextEndBrace));        
        return wholeDTS.indexOf(";", nextEndBrace);
    }
}