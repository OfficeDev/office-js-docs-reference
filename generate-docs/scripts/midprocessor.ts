import * as fsx from 'fs-extra';
import * as path from "path";
// import yaml = require('js-yaml');

// enum Host {
//     common,
//     excel,
//     onenote,
//     outlook,
//     powerpoint,
//     visio,
//     word
// }

// const CURRENT_EXCEL_RELEASE = 8;
// const CURRENT_OUTLOOK_RELEASE = 7;
// const CURRENT_WORD_RELEASE = 3;

tryCatch(async () => {
    console.log("Copying snippets file to subfolders");
    const snippets = path.resolve("../json/snippets.yaml");
    // let allSnippets: Object = yaml.safeLoad(fsx.readFileSync(snippetPath).toString());
    // let excelSnippets: Object = yaml.safeLoad("");
    // for (const key of Object.keys(allSnippets)) {
    //     if (key.startsWith("Excel")) {
    //         inTheseDtses(Host.excel, key); // TODO
    //         //excelSnippets += yaml.safeLoad(allSnippets[key]);
    //     }
    // }

    //console.log(excelSnippets);

    fsx.copySync(snippets, "../json/excel/snippets.yaml");
    fsx.copySync(snippets, "../json/excel_1_1/snippets.yaml");
    fsx.copySync(snippets, "../json/excel_1_2/snippets.yaml");
    fsx.copySync(snippets, "../json/excel_1_3/snippets.yaml");
    fsx.copySync(snippets, "../json/excel_1_4/snippets.yaml");
    fsx.copySync(snippets, "../json/excel_1_5/snippets.yaml");
    fsx.copySync(snippets, "../json/excel_1_6/snippets.yaml");
    fsx.copySync(snippets, "../json/excel_1_7/snippets.yaml");
    fsx.copySync(snippets, "../json/excel_1_8/snippets.yaml");

    fsx.copySync(snippets, "../json/office/snippets.yaml");

    fsx.copySync(snippets, "../json/onenote/snippets.yaml");

    fsx.copySync(snippets, "../json/outlook/snippets.yaml");
    fsx.copySync(snippets, "../json/outlook_1_1/snippets.yaml");
    fsx.copySync(snippets, "../json/outlook_1_2/snippets.yaml");
    fsx.copySync(snippets, "../json/outlook_1_3/snippets.yaml");
    fsx.copySync(snippets, "../json/outlook_1_4/snippets.yaml");
    fsx.copySync(snippets, "../json/outlook_1_5/snippets.yaml");
    fsx.copySync(snippets, "../json/outlook_1_6/snippets.yaml");
    fsx.copySync(snippets, "../json/outlook_1_7/snippets.yaml");

    fsx.copySync(snippets, "../json/powerpoint/snippets.yaml");

    fsx.copySync(snippets, "../json/visio/snippets.yaml");

    fsx.copySync(snippets, "../json/word/snippets.yaml");
    fsx.copySync(snippets, "../json/word_1_1/snippets.yaml");
    fsx.copySync(snippets, "../json/word_1_2/snippets.yaml");
    fsx.copySync(snippets, "../json/word_1_3/snippets.yaml");

    console.log("Moving Custom Functions APIs to correct versions of Excel");
    const customFunctionsJson = path.resolve("../json/custom-functions-runtime.api.json");
    const officeRuntimeJson = path.resolve("../json/office-runtime.api.json");
    fsx.copySync(customFunctionsJson, "../json/excel/custom-functions-runtime.api.json");
    fsx.copySync(officeRuntimeJson, "../json/excel/office-runtime.api.json");
});

// function inTheseDtses(hostName: Host, fieldName: string): string[] {
//     let dtsList = [];
//     switch (hostName) {
//         case Host.common:
//             const common = path.resolve("../api-extractor-inputs-common/common.d.ts");
//             if (fsx.readFileSync(common).toString().includes(fieldName)) {
//                 dtsList.push(common);
//             } else {
//                 console.warn("Missing field for snippet: " + fieldName);
//             }
//             break;
//         case Host.excel:
//             const excelPreview = path.resolve("../api-extractor-inputs-excel/excel.d.ts");
//             if (fsx.readFileSync(excelPreview).toString().includes(fieldName)) {
//                 dtsList.push(excelPreview);
//                 for (let i = CURRENT_EXCEL_RELEASE; i > 0; i--) {
//                     const currentVersionPath = path.resolve(`../api-extractor-inputs-excel-release/excel_1_${i}/excel.d.ts`);
//                     if (fsx.readFileSync(currentVersionPath).toString().includes(fieldName)) {
//                         dtsList.push(currentVersionPath);
//                     } else {
//                         break; // it's not going to be in older versions
//                     }
//                 }
//             } else {
//                 console.warn("Missing field for snippet: " + fieldName);
//             }
//             break;
//         case Host.onenote:
//             const onenote = path.resolve("../api-extractor-inputs-onenote/onenote.d.ts");
//             if (fsx.readFileSync(onenote).toString().includes(fieldName)) {
//                 dtsList.push(onenote);
//             } else {
//                 console.warn("Missing field for snippet: " + fieldName);
//             }
//             break;
//         case Host.outlook:
//             const outlookPreview = path.resolve("../api-extractor-inputs-outlook/outlook.d.ts");
//             if (fsx.readFileSync(outlookPreview).toString().includes(fieldName)) {
//                 dtsList.push(outlookPreview);
//                 for (let i = CURRENT_OUTLOOK_RELEASE; i > 0; i--) {
//                     const currentVersionPath = path.resolve(`../api-extractor-inputs-outlook-release/outlook_1_${i}/outlook.d.ts`);
//                     if (fsx.readFileSync(currentVersionPath).toString().includes(fieldName)) {
//                         dtsList.push(currentVersionPath);
//                     } else {
//                         break; // it's not going to be in older versions
//                     }
//                 }
//             } else {
//                 console.warn("Missing field for snippet: " + fieldName);
//             }
//             break;
//         case Host.powerpoint:
//             const powerpoint = path.resolve("../api-extractor-inputs-powerpoint/powerpoint.d.ts");
//             if (fsx.readFileSync(powerpoint).toString().includes(fieldName)) {
//                 dtsList.push(powerpoint);
//             } else {
//                 console.warn("Missing field for snippet: " + fieldName);
//             }
//             break;
//         case Host.visio:
//             const visio = path.resolve("../api-extractor-inputs-visio/visio.d.ts");
//             if (fsx.readFileSync(visio).toString().includes(fieldName)) {
//                 dtsList.push(visio);
//             } else {
//                 console.warn("Missing field for snippet: " + fieldName);
//             }
//             break;
//         case Host.word:
//             const wordPreview = path.resolve("../api-extractor-inputs-word/word.d.ts");
//             if (fsx.readFileSync(wordPreview).toString().includes(fieldName)) {
//                 dtsList.push(wordPreview);
//                 for (let i = CURRENT_WORD_RELEASE; i > 0; i--) {
//                     const currentVersionPath = path.resolve(`../api-extractor-inputs-outlook-release/word_1_${i}/word.d.ts`);
//                     if (fsx.readFileSync(currentVersionPath).toString().includes(fieldName)) {
//                         dtsList.push(currentVersionPath);
//                     } else {
//                         break; // it's not going to be in older versions
//                     }
//                 }
//             } else {
//                 console.warn("Missing field for snippet: " + fieldName);
//             }
//             break;
//     }

//     return dtsList;
// }

async function tryCatch(call: () => Promise<void>) {
    try {
        await call();
    } catch (e) {
        console.error(e);
        process.exit(1);
    }
}
