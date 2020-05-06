import { fetchAndThrowOnError } from './util';
import * as fsx from 'fs-extra';
import * as path from "path";
import yaml = require('js-yaml');
import * as colors from 'colors';

const CURRENT_EXCEL_RELEASE = 11;
const OLDEST_EXCEL_RELEASE_WITH_CUSTOM_FUNCTIONS = 9;
const CURRENT_OUTLOOK_RELEASE = 8;
const CURRENT_WORD_RELEASE = 3;

tryCatch(async () => {
    // ----
    // Clean up Office and Outlook json cross-referencing.
    // ----
    console.log("\nCleaning up Office json cross-referencing...");

    const officeJsonPath = path.resolve("../json/office");
    const officeFilename = "office.api.json";
    fsx.writeFileSync(officeJsonPath + '/' + officeFilename, fsx.readFileSync(officeJsonPath + '/' + officeFilename).toString().replace("office!Office.Mailbox", "outlook!Office.Mailbox").replace("office!Office.RoamingSettings", "outlook!Office.RoamingSettings"));

    console.log("\nCompleted Office json cross-referencing cleanup");

    console.log("\nCleaning up Outlook json cross-referencing...");

    const outlookJsonPath = path.resolve("../json/outlook");
    const outlookFilename = "outlook.api.json";
    console.log("\nStarting outlook...");
    let outlookJson = fsx.readFileSync(`${outlookJsonPath}/${outlookFilename}`).toString();
    fsx.writeFileSync(`${outlookJsonPath}/${outlookFilename}`, cleanUpOutlookJson(outlookJson));
    console.log("\Completed outlook");
    for (let i = CURRENT_OUTLOOK_RELEASE; i > 0; i--) {
        console.log(`\nStarting outlook_1_${i}...`);
        outlookJson = fsx.readFileSync(`${outlookJsonPath}_1_${i}/${outlookFilename}`).toString();
        fsx.writeFileSync(`${outlookJsonPath}_1_${i}/${outlookFilename}`, cleanUpOutlookJson(outlookJson));
        console.log(`Completed outlook_1_${i}`);
    }

    console.log("\nCompleted Outlook json cross-referencing cleanup");

    console.log("\nCleaning up Excel json cross-referencing...");

    const excelJsonPath = path.resolve("../json/excel");
    const excelFilename = "excel.api.json";
    console.log("\nStarting excel...");
    let excelJson = fsx.readFileSync(`${excelJsonPath}/${excelFilename}`).toString();
    fsx.writeFileSync(`${excelJsonPath}/${excelFilename}`, cleanUpRichApiJson(excelJson));
    excelJson = fsx.readFileSync(`${excelJsonPath}_online/${excelFilename}`).toString();
    fsx.writeFileSync(`${excelJsonPath}_online/${excelFilename}`, cleanUpRichApiJson(excelJson));
    console.log("Completed excel");
    for (let i = CURRENT_EXCEL_RELEASE; i > 0; i--) {
        console.log(`\nStarting excel${i}...`);
        excelJson = fsx.readFileSync(`${excelJsonPath}_1_${i}/${excelFilename}`).toString();
        fsx.writeFileSync(`${excelJsonPath}_1_${i}/${excelFilename}`, cleanUpRichApiJson(excelJson));
        console.log(`Completed excel${i}`);
    }

    console.log("\nCompleted Excel json cross-referencing cleanup");

    console.log("\nCleaning up Word json cross-referencing...");

    const wordJsonPath = path.resolve("../json/word");
    const wordFilename = "word.api.json";
    console.log("\nStarting word...");
    let wordJson = fsx.readFileSync(`${wordJsonPath}/${wordFilename}`).toString();
    fsx.writeFileSync(`${wordJsonPath}/${wordFilename}`, cleanUpRichApiJson(wordJson));
    console.log("Completed word");
    for (let i = CURRENT_WORD_RELEASE; i > 0; i--) {
        console.log(`\nStarting word_1_${i}...`);
        wordJson = fsx.readFileSync(`${wordJsonPath}_1_${i}/${wordFilename}`).toString();
        fsx.writeFileSync(`${wordJsonPath}_1_${i}/${wordFilename}`, cleanUpRichApiJson(wordJson));
        console.log(`Completed word_1_${i}`);
    }

    console.log("\nCompleted Word json cross-referencing cleanup");

    console.log("\nCleaning up Visio json cross-referencing...");

    const visioJsonPath = path.resolve("../json/visio");
    const visioFilename = "visio.api.json";
    console.log("\nStarting visio...");
    let visioJson = fsx.readFileSync(`${visioJsonPath}/${visioFilename}`).toString();
    fsx.writeFileSync(`${visioJsonPath}/${visioFilename}`, cleanUpRichApiJson(visioJson));
    console.log("Completed visio");

    console.log("\nCompleted Visio json cross-referencing cleanup");

    console.log("\nCleaning up OneNote json cross-referencing...");

    const onenoteJsonPath = path.resolve("../json/onenote");
    const onenoteFilename = "onenote.api.json";
    console.log("\nStarting onenote...");
    let onenoteJson = fsx.readFileSync(`${onenoteJsonPath}/${onenoteFilename}`).toString();
    fsx.writeFileSync(`${onenoteJsonPath}/${onenoteFilename}`, cleanUpRichApiJson(onenoteJson));
    console.log("Completed onenote");
    console.log("\nCompleted OneNote json cross-referencing cleanup");

    // ----
    // Process Snippets
    // ----
    console.log("\nRemoving old snippets input files...");

    const scriptInputsPath = path.resolve("../script-inputs");
    fsx.readdirSync(scriptInputsPath)
        .filter(filename => filename.indexOf("snippets") > 0)
        .forEach(filename => fsx.removeSync(scriptInputsPath + '/' + filename));

    console.log("\nCreating snippets file...");

    console.log("\nReading from: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/master/snippet-extractor-output/snippets.yaml");
    fsx.writeFileSync("../script-inputs/script-lab-snippets.yaml", await fetchAndThrowOnError("https://raw.githubusercontent.com/OfficeDev/office-js-snippets/master/snippet-extractor-output/snippets.yaml", "text"));

    console.log("\nReading from files: " + path.resolve("../../docs/code-snippets"));

    const snippetsSourcePath = path.resolve("../../docs/code-snippets");
    let localCodeSnippetsString : string = "";
    fsx.readdirSync(path.resolve(snippetsSourcePath))
        .filter(name => name.endsWith('.yaml') || name.endsWith('.yml'))
        .forEach((filename, index) => {
            localCodeSnippetsString += fsx.readFileSync(`${snippetsSourcePath}/${filename}`).toString() + "\r\n";
        });

    fsx.writeFileSync("../script-inputs/local-repo-snippets.yaml", localCodeSnippetsString);

    // Parse the YAML into an object/hash set.
    let allSnippets: Object = yaml.load(localCodeSnippetsString);

    // If a duplicate key exists, merge the Script Lab example(s) into the item with the existing key.
    let scriptLabSnippets: Object = yaml.load(fsx.readFileSync(`../script-inputs/script-lab-snippets.yaml`).toString());
    for (const key of Object.keys(scriptLabSnippets)) {
        if (allSnippets[key]) {
            console.log("Combining local and Script Lab snippets for: " + key);
            allSnippets[key] = allSnippets[key].concat(scriptLabSnippets[key]);
        } else {
            allSnippets[key] = scriptLabSnippets[key];
        }
    }

    console.log("\nWriting snippets to: " + path.resolve("../json/snippets.yaml"));

    fsx.writeFileSync("../json/snippets.yaml", yaml.safeDump(
        allSnippets,
        {sortKeys: <any>((a: string, b: string) => {
            if (a < b) {
                return -1;
            } else if (a > b) {
                return 1;
            } else {
                return 0;
            }
        })}
    ));

    console.log("Copying snippets file to subfolders");
    const snippetPath = path.resolve("../json/snippets.yaml");
    allSnippets = yaml.safeLoad(fsx.readFileSync(snippetPath).toString(), {strict: true});
    let commonSnippetKeys = [];
    let excelSnippetKeys = [];
    let onenoteSnippetKeys = [];
    let outlookSnippetKeys = [];
    let powerpointSnippetKeys = [];
    let visioSnippetKeys = [];
    let wordSnippetKeys = [];
    let commonText = fsx.readFileSync(path.resolve("../json/office/office.api.json"));
    for (const key of Object.keys(allSnippets)) {
        if (key.startsWith("Excel")) {
            excelSnippetKeys.push(key);
        } else if (key.startsWith("OneNote")) {
            onenoteSnippetKeys.push(key);
        } else if (key.startsWith("PowerPoint")) {
            powerpointSnippetKeys.push(key);
        } else if (key.startsWith("Visio")) {
            visioSnippetKeys.push(key);
        } else if (key.startsWith("Word")) {
            wordSnippetKeys.push(key);
        } else if (key.startsWith("Office")) {
            if (commonText.indexOf(key) >= 0) {
                commonSnippetKeys.push(key);
            } else {
                outlookSnippetKeys.push(key);
            }
        } else {
            console.error(colors.red("Unknown snippet key prefix: " + key));
        }
    }

    let commonSnippets = {};
    let excelSnippets = {};
    let onenoteSnippets = {};
    let outlookSnippets = {};
    let powerpointSnippets = {};
    let visioSnippets = {};
    let wordSnippets = {};

    commonSnippetKeys.forEach(key => {
        commonSnippets[key] = allSnippets[key];
        delete allSnippets[key];
    });
    excelSnippetKeys.forEach(key => {
        excelSnippets[key] = allSnippets[key];
        delete allSnippets[key];
    });
    onenoteSnippetKeys.forEach(key => {
        onenoteSnippets[key] = allSnippets[key];
        delete allSnippets[key];
    });
    outlookSnippetKeys.forEach(key => {
        outlookSnippets[key] = allSnippets[key];
        delete allSnippets[key];
    });
    powerpointSnippetKeys.forEach(key => {
        powerpointSnippets[key] = allSnippets[key];
        delete allSnippets[key];
    });
    visioSnippetKeys.forEach(key => {
        visioSnippets[key] = allSnippets[key];
        delete allSnippets[key];
    });
    wordSnippetKeys.forEach(key => {
        wordSnippets[key] = allSnippets[key];
        delete allSnippets[key];
    });

    fsx.writeFileSync("../json/excel/snippets.yaml", yaml.safeDump(excelSnippets));
    fsx.writeFileSync("../json/excel_online/snippets.yaml", yaml.safeDump(excelSnippets));
    for (let i = CURRENT_EXCEL_RELEASE; i > 0; i--) {
        fsx.writeFileSync(`../json/excel_1_${i}/snippets.yaml`, yaml.safeDump(excelSnippets));
    }

    fsx.writeFileSync("../json/office/snippets.yaml", yaml.safeDump(commonSnippets));

    fsx.writeFileSync("../json/onenote/snippets.yaml", yaml.safeDump(onenoteSnippets));

    fsx.writeFileSync("../json/outlook/snippets.yaml", yaml.safeDump(outlookSnippets));
    for (let i = CURRENT_OUTLOOK_RELEASE; i > 0; i--) {
        fsx.writeFileSync(`../json/outlook_1_${i}/snippets.yaml`, yaml.safeDump(outlookSnippets));
    }

    fsx.writeFileSync("../json/powerpoint/snippets.yaml", yaml.safeDump(powerpointSnippets));

    fsx.writeFileSync("../json/visio/snippets.yaml", yaml.safeDump(visioSnippets));

    fsx.writeFileSync("../json/word/snippets.yaml", yaml.safeDump(wordSnippets));
    for (let i = CURRENT_WORD_RELEASE; i > 0; i--) {
        fsx.writeFileSync(`../json/word_1_${i}/snippets.yaml`, yaml.safeDump(wordSnippets));
    }

    console.log("Moving Custom Functions APIs to correct versions of Excel");
    const customFunctionsJson = path.resolve("../json/custom-functions-runtime.api.json");
    const officeRuntimeJson = path.resolve("../json/office-runtime.api.json");
    fsx.copySync(customFunctionsJson, "../json/excel/custom-functions-runtime.api.json");
    for (let i = CURRENT_EXCEL_RELEASE; i >= OLDEST_EXCEL_RELEASE_WITH_CUSTOM_FUNCTIONS; i--) {
        fsx.copySync(customFunctionsJson, `../json/excel_1_${i}/custom-functions-runtime.api.json`);
    }
    fsx.copySync(customFunctionsJson, `../json/excel_online/custom-functions-runtime.api.json`);

    console.log("Moving Office Runtime APIs to Common API");
    fsx.copySync(officeRuntimeJson, `../json/office/office-runtime.api.json`);
});

function cleanUpOutlookJson(jsonString : string) {
    const outlookSearchString = "outlook!";
    const commonApiSearchString = "CommonAPI";
    let startSearchIndex = jsonString.indexOf(commonApiSearchString);
    do {
        let outlookIndex = jsonString.indexOf(outlookSearchString, startSearchIndex);
        jsonString = jsonString.substring(0, outlookIndex) + "office!" + jsonString.substring(outlookIndex + 8);
        startSearchIndex = jsonString.indexOf(commonApiSearchString, outlookIndex + 8);
    } while (startSearchIndex >= 0);
    return jsonString;
}

function cleanUpRichApiJson(jsonString : string) {
    return jsonString.replace(/(excel|word|visio|onenote)\!OfficeExtension/g, "office!OfficeExtension");
}

async function tryCatch(call: () => Promise<void>) {
    try {
        await call();
    } catch (e) {
        console.error(e);
        process.exit(1);
    }
}
