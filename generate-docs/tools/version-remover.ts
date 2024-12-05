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
    if (declarationString.indexOf("class") >= 0 || 
        declarationString.indexOf("enum") >= 0 || 
        declarationString.indexOf("interface") >= 0) {
        // Discount internal bracket pairs.
        let nextStartBrace = wholeDts.indexOf("{", endPosition);
        let nextEndBrace = wholeDts.indexOf("}", endPosition);
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
    wholeDts = wholeDts.replace(/content: CommentRichContent \| string,/g, "content: string,");
}

if (process.argv[3] === "Mailbox 1.14") {
    console.log("Address SpamReportingEventCompletedOptions reference when removing Mailbox 1.14");
    wholeDts = wholeDts.replace(/options\?: SmartAlertsEventCompletedOptions \| SpamReportingEventCompletedOptions/g, "options?: SmartAlertsEventCompletedOptions");
}

if (process.argv[3] === "Mailbox 1.12") {
    console.log("Address SmartAlertsEventCompletedOptions reference when removing Mailbox 1.12");
    wholeDts = wholeDts.replace(/options\?: SmartAlertsEventCompletedOptions/g, "");
    wholeDts = wholeDts.replace(/@param options - Optional. An object that specifies the behavior of an event-based or spam-reporting add-in when it completes processing an event./g, "");
}

if (process.argv[3] === "WordApiDesktop 1.1") {
    console.log("Address ImportedStylesConflictBehavior reference when removing WordApiDesktop 1.1");
    wholeDts = wholeDts.replace(/importedStylesConflictBehavior\?: Word.ImportedStylesConflictBehavior/g, "importedStylesConflictBehavior?: Word.ImportedStylesConflictBehavior");
}

fsx.writeFileSync(process.argv[4], wholeDts);


function getDeclarationEnd(wholeDts: string, startIndex: number): number {
    let nextSemicolon = wholeDts.indexOf(";", startIndex);
    let nextNewLine = wholeDts.indexOf("\n", startIndex);
    let nextStartBrace = wholeDts.indexOf("{", startIndex);
    let nextEndBrace = wholeDts.indexOf("}", startIndex);
    let nextComment = wholeDts.indexOf("/**", startIndex);
    
    // Figure out if the declaration has an internal class.
    if (nextSemicolon < nextNewLine) {
        return nextSemicolon;
    } else if (nextStartBrace > nextEndBrace && nextEndBrace < nextComment) {
        return wholeDts.lastIndexOf("\n", nextEndBrace);
    } else {
        return wholeDts.lastIndexOf("\n", nextComment);
    }
}