import * as fsx from "fs-extra";

if (process.argv.length !== 5 || process.argv.find((x: string) => {return x === "-?"})) {
    console.log("usage: node version-remover [source d.ts] [API set name] [output file name]");
    console.log("example: node version-remover excel.d.ts \"ExcelApi 1.8\" excel_1_7.d.ts");
    process.exit(0);
}

console.log("Version Remover - Creating " + process.argv[4]);
let wholeDts = fsx.readFileSync(process.argv[2]).toString();
let declarationString;
// find the API tag
let indexOfApiSetTag = wholeDts.indexOf("Api set: " + process.argv[3]);
while (indexOfApiSetTag >= 0) {
    // find the comment block around the API tag
    let commentStart = wholeDts.lastIndexOf("/**", indexOfApiSetTag);
    let commentEnd = wholeDts.indexOf("*/", indexOfApiSetTag);
    commentEnd =  wholeDts.indexOf("\n", commentEnd) + 1; // Account for newline and ending characters.

    // the declaration string is the line following the comment
    declarationString = wholeDts.substring(commentEnd, wholeDts.indexOf("\n", commentEnd));
    let endPosition = commentEnd + declarationString.length;
    if (declarationString.indexOf("class") >= 0 || declarationString.indexOf("enum") >= 0 || declarationString.indexOf("interface") >= 0) {
        let nextStartBrace = wholeDts.indexOf("{", endPosition);
        let nextEndBrace = wholeDts.indexOf("}", endPosition);
        // Discount internal bracket pairs.
        while (nextStartBrace < nextEndBrace && nextStartBrace >= 0) {
            nextStartBrace = wholeDts.indexOf("{", nextStartBrace + 1);
            nextEndBrace = wholeDts.indexOf("}", nextEndBrace + 1);
        }
        endPosition = wholeDts.indexOf("}", nextEndBrace - 1);
    } else {
        endPosition = getDeclarationEnd(wholeDts, commentEnd);
    }

    if (endPosition === -1) {
        endPosition = commentEnd;
    }
    wholeDts = wholeDts.substring(0, commentStart) + wholeDts.substring(endPosition + 1);
    indexOfApiSetTag = wholeDts.indexOf("Api set: " + process.argv[3]);
}

/* Add necessary custom logic here*/

if (process.argv[3] === "ExcelApi 1.11") {
    console.log("Address CommentRichContent reference for when removing ExcelApi 1.11");
    wholeDts = wholeDts.replace(/add\(content: CommentRichContent \| string,/g, "add(content: string,").
                replace(/add\(cellAddress: Range \| string, content: CommentRichContent \| string,/g, "add(cellAddress: Range | string, content: string,");
}

fsx.writeFileSync(process.argv[4], wholeDts);


function getDeclarationEnd(wholeDts: string, startIndex: number): number {
    let nextSemicolon = wholeDts.indexOf(";", startIndex);
    let nextNewLine = wholeDts.indexOf("\n", startIndex);
    let nextStartBrace = wholeDts.indexOf("{", startIndex);
    let nextEndBrace = wholeDts.indexOf("}", startIndex);
    
    // Figure out if the declaration has an internal class.
    if (nextSemicolon < nextNewLine) {
        return nextSemicolon;
    } else if (nextStartBrace > nextNewLine) {
        // No semicolon or braces means this is an enum member.
        return nextNewLine;
    } else {
        // The declaration is on multiple lines, likely due to an internal class.
        wholeDts.substring(startIndex, wholeDts.indexOf(";", nextEndBrace));        
        return wholeDts.indexOf(";", nextEndBrace);
    }
}