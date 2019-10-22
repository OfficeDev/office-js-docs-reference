import * as fsx from 'fs-extra';
import * as path from "path";
import yaml = require('js-yaml');
import * as colors from 'colors';

const CURRENT_EXCEL_RELEASE = 10;
const OLDEST_EXCEL_RELEASE_WITH_CUSTOM_FUNCTIONS = 9;
const CURRENT_OUTLOOK_RELEASE = 7;
const CURRENT_WORD_RELEASE = 3;

tryCatch(async () => {
    console.log("Copying snippets file to subfolders");
    const snippets = path.resolve("../json/snippets.yaml");
    let allSnippets: Object = yaml.safeLoad(fsx.readFileSync(snippets).toString(), {strict: true});
    let excelSnippetKeys = [];
    let onenoteSnippetKeys = [];
    let outlookAndCommonSnippetKeys = [];
    let powerpointSnippetKeys = [];
    let visioSnippetKeys = [];
    let wordSnippetKeys = [];
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
            outlookAndCommonSnippetKeys.push(key);
        } else {
            console.error(colors.red("Unknown snippet key prefix: " + key));
        }
    }

    let excelSnippets = {};
    let onenoteSnippets = {};
    let outlookAndCommonSnippets = {};
    let powerpointSnippets = {};
    let visioSnippets = {};
    let wordSnippets = {};

    excelSnippetKeys.forEach(key => {
        excelSnippets[key] = allSnippets[key];
        delete allSnippets[key];
    });
    onenoteSnippetKeys.forEach(key => {
        onenoteSnippets[key] = allSnippets[key];
        delete allSnippets[key];
    });
    outlookAndCommonSnippetKeys.forEach(key => {
        outlookAndCommonSnippets[key] = allSnippets[key];
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

    fsx.writeFileSync("../json/office/snippets.yaml", yaml.safeDump(outlookAndCommonSnippets));

    fsx.writeFileSync("../json/onenote/snippets.yaml", yaml.safeDump(onenoteSnippets));

    fsx.writeFileSync("../json/outlook/snippets.yaml", yaml.safeDump(outlookAndCommonSnippets));
    for (let i = CURRENT_OUTLOOK_RELEASE; i > 0; i--) {
        fsx.writeFileSync(`../json/outlook_1_${i}/snippets.yaml`, yaml.safeDump(outlookAndCommonSnippets));
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
    fsx.copySync(officeRuntimeJson, "../json/excel/office-runtime.api.json");
    for (let i = CURRENT_EXCEL_RELEASE; i >= OLDEST_EXCEL_RELEASE_WITH_CUSTOM_FUNCTIONS; i--) {
        fsx.copySync(customFunctionsJson, `../json/excel_1_${i}/custom-functions-runtime.api.json`);
        fsx.copySync(officeRuntimeJson, `../json/excel_1_${i}/office-runtime.api.json`);
    }
});

async function tryCatch(call: () => Promise<void>) {
    try {
        await call();
    } catch (e) {
        console.error(e);
        process.exit(1);
    }
}
