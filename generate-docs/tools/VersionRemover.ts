// usage: node VersionRemover [source d.ts] [API set name] [output file name]
// example: node VersionRemover excel.d.ts "ExcelApi 1.8" excel_1_7.d.ts
import * as fsx from "fs-extra";

let wholeDTS = fsx.readFileSync(process.argv[2]).toString();
let indexOfApiSetTag = wholeDTS.indexOf("Api set: " + process.argv[3]);
while (indexOfApiSetTag >= 0) {
    let commentStart = wholeDTS.lastIndexOf("/**", indexOfApiSetTag);
    let commentEnd = wholeDTS.indexOf("*/", indexOfApiSetTag) + 1;
    let declarationString = wholeDTS.substring(commentEnd + 1, wholeDTS.indexOf("\n", commentEnd + 2));
    let endPosition;
    if (declarationString.indexOf("class") >= 0 || declarationString.indexOf("enum") >= 0 || declarationString.indexOf("interface") >= 0) {
        endPosition = wholeDTS.indexOf("}\n", commentEnd);
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
