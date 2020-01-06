import * as fsx from "fs-extra";

if (process.argv.length != 5 || process.argv.find((x: string) => {return x === "-?"})) {
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
    let endPosition;
    if (declarationString.indexOf("class") >= 0 || declarationString.indexOf("enum") >= 0 || declarationString.indexOf("interface") >= 0) {
        endPosition = Math.max(wholeDTS.indexOf("}\r\n", commentEnd), wholeDTS.indexOf("}\n", commentEnd));
    } else {
        endPosition = wholeDTS.indexOf(";", commentEnd);
    }

    if (endPosition === -1) {
        endPosition = commentEnd;
    }
    wholeDTS = wholeDTS.substring(0, commentStart) + wholeDTS.substring(endPosition + 1);
    indexOfApiSetTag = wholeDTS.indexOf("Api set: " + process.argv[3]);
}

fsx.writeFileSync(process.argv[4], wholeDTS);
