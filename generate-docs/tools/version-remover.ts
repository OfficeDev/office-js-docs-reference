import * as fsx from "fs-extra";

if (process.argv.length !== 5 || process.argv.find((x: string) => {return x === "-?"})) {
    console.log("usage: node version-remover [source d.ts] [API set name] [output file name]");
    console.log("example: node version-remover excel.d.ts \"ExcelApi 1.8\" excel_1_7.d.ts");
    process.exit(0);
}

console.log("Version Remover - Creating " + process.argv[4]);
let wholeDTS = fsx.readFileSync(process.argv[2]).toString();

// find the API tag
let indexOfApiSetTag = wholeDTS.indexOf("Api set: " + process.argv[3]);
while (indexOfApiSetTag >= 0) {
    // find the comment block around the API tag
    let commentStart = wholeDTS.lastIndexOf("/**", indexOfApiSetTag);
    let commentEnd = wholeDTS.indexOf("*/", indexOfApiSetTag) + 3; // +3 to include the ending characters and newline

    // the declaration string is the line following the comment
    let declarationString = wholeDTS.substring(commentEnd, wholeDTS.indexOf("\n", commentEnd + 2));
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
    let nextEndBrace = wholeDTS.indexOf("}", startIndex);
    
    if (nextSemicolon < nextNewLine) {
        return nextSemicolon;
    } else {
        // The declaration is on multiple lines, likely due to an internal class.
        wholeDTS.substring(startIndex, wholeDTS.indexOf(";", nextEndBrace));        
        return wholeDTS.indexOf(";", nextEndBrace);
    }
}