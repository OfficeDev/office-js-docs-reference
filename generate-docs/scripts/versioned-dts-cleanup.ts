// usage: node versioned-dts-cleanup [source d.ts] [host name] [version number]
// example: node versioned-dts-cleanup outlook.d.ts Outlook 1.6
import * as fsx from "fs-extra";

console.log(`Cleaning up ${process.argv[3]} d.ts file (${process.argv[2]}`);
let wholeDTS = fsx.readFileSync(process.argv[2]).toString();

if (process.argv[3] === "Outlook") {
    wholeDTS = wholeDTS.replace(/\/objectmodel\/requirement-set-1.[\d]*\//g, `/objectmodel/requirement-set-${process.argv[4]}/`);
} else {
    console.log(`Host ${process.argv[3]} has no defined clean up steps.`);
}

fsx.writeFileSync(process.argv[2], wholeDTS);
