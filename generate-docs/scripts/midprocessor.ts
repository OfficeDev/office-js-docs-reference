import { fetchAndThrowOnError } from './util';
import * as fsx from 'fs-extra';
import * as path from "path";
import yaml = require('js-yaml');
import * as colors from 'colors';

const CURRENT_EXCEL_RELEASE = 15;
const OLDEST_EXCEL_RELEASE_WITH_CUSTOM_FUNCTIONS = 9;
const CURRENT_OUTLOOK_RELEASE = 12;
const CURRENT_WORD_RELEASE = 4;
const CURRENT_POWERPOINT_RELEASE = 4;

tryCatch(async () => {
    // ----
    // Clean up Office and Outlook json cross-referencing.
    // ----
    console.log("\nCleaning up Office json cross-referencing...");

    const officeJsonPaths: string[] = [path.resolve("../json/office"), path.resolve("../json/office_release")];
    const officeFilename = "office.api.json";
    officeJsonPaths.forEach((officeJsonPath) => {
        fsx.writeFileSync(
            officeJsonPath + '/' + officeFilename,
            fsx.readFileSync(officeJsonPath + '/' + officeFilename)
                .toString()
                .replace(/office\!Office\.Mailbox/g, "outlook!Office.Mailbox")
                .replace(/office\!Office\.RoamingSettings/g, "outlook!Office.RoamingSettings"));
    });

    console.log("\nCompleted Office json cross-referencing cleanup");

    cleanUpJson("outlook");
    cleanUpJson("excel");
    cleanUpJson("word");
    cleanUpJson("powerpoint");
    cleanUpJson("onenote");
    cleanUpJson("visio");

    // ----
    // Process Snippets
    // ----
    console.log("\nRemoving old snippets input files...");

    const scriptInputsPath = path.resolve("../script-inputs");
    fsx.readdirSync(scriptInputsPath)
        .filter(filename => filename.indexOf("snippets") > 0)
        .forEach(filename => fsx.removeSync(scriptInputsPath + '/' + filename));

    console.log("\nCreating snippets file...");

    console.log("\nReading from: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/snippet-extractor-output/snippets.yaml");
    fsx.writeFileSync("../script-inputs/script-lab-snippets.yaml", await fetchAndThrowOnError("https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/snippet-extractor-output/snippets.yaml", "text"));

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
        if (key.startsWith("Excel") || key.startsWith("CustomFunctions")) {
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

    writeSnippetFileAndClearYamlIfNew("../json/excel/snippets.yaml", yaml.safeDump(excelSnippets), "excel");
    writeSnippetFileAndClearYamlIfNew("../json/excel_online/snippets.yaml", yaml.safeDump(excelSnippets), "excel");
    for (let i = CURRENT_EXCEL_RELEASE; i > 0; i--) {
        writeSnippetFileAndClearYamlIfNew(`../json/excel_1_${i}/snippets.yaml`, yaml.safeDump(excelSnippets), "excel");
    }

    writeSnippetFileAndClearYamlIfNew("../json/office/snippets.yaml", yaml.safeDump(commonSnippets), "office");
    writeSnippetFileAndClearYamlIfNew("../json/office_release/snippets.yaml", yaml.safeDump(commonSnippets), "office");

    writeSnippetFileAndClearYamlIfNew("../json/onenote/snippets.yaml", yaml.safeDump(onenoteSnippets), "onenote");

    writeSnippetFileAndClearYamlIfNew("../json/outlook/snippets.yaml", yaml.safeDump(outlookSnippets), "outlook");
    for (let i = CURRENT_OUTLOOK_RELEASE; i > 0; i--) {
        writeSnippetFileAndClearYamlIfNew(`../json/outlook_1_${i}/snippets.yaml`, yaml.safeDump(outlookSnippets), "outlook");
    }

    writeSnippetFileAndClearYamlIfNew("../json/powerpoint/snippets.yaml", yaml.safeDump(powerpointSnippets), "powerpoint");
    for (let i = CURRENT_POWERPOINT_RELEASE; i > 0; i--) {
        writeSnippetFileAndClearYamlIfNew(`../json/powerpoint_1_${i}/snippets.yaml`, yaml.safeDump(powerpointSnippets), "powerpoint");
    }

    writeSnippetFileAndClearYamlIfNew("../json/visio/snippets.yaml", yaml.safeDump(visioSnippets), "visio");

    writeSnippetFileAndClearYamlIfNew("../json/word/snippets.yaml", yaml.safeDump(wordSnippets), "word");
    writeSnippetFileAndClearYamlIfNew("../json/word_online/snippets.yaml", yaml.safeDump(wordSnippets), "word");
    writeSnippetFileAndClearYamlIfNew("../json/word_1_4_hidden_document/snippets.yaml", yaml.safeDump(wordSnippets), "word");
    writeSnippetFileAndClearYamlIfNew("../json/word_1_3_hidden_document/snippets.yaml", yaml.safeDump(wordSnippets), "word");
    for (let i = CURRENT_WORD_RELEASE; i > 0; i--) {
        writeSnippetFileAndClearYamlIfNew(`../json/word_1_${i}/snippets.yaml`, yaml.safeDump(wordSnippets), "word");
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
    fsx.copySync(officeRuntimeJson, `../json/office_release/office-runtime.api.json`);

    console.log("Cleaning up What's New markdown files.");
    let filePath = `../../docs/includes/outlook-preview.md`;
    fsx.writeFileSync(filePath, cleanUpOutlookMarkdown(fsx.readFileSync(filePath).toString()));
    for (let i = CURRENT_OUTLOOK_RELEASE; i > 0; i--) {
        filePath = `../../docs/includes/outlook-1_${i}.md`;
        fsx.writeFileSync(filePath, cleanUpOutlookMarkdown(fsx.readFileSync(filePath).toString()));
    }
});

function cleanUpJson(host: string) {
    console.log(`\nCleaning up ${host} json cross-referencing...`);

    const jsonPath = path.resolve(`../json/${host}`);
    const fileName = `${host}.api.json`;
    console.log(`\nStarting ${host}...`);
    let json = fsx.readFileSync(`${jsonPath}/${fileName}`).toString();
    let cleanJson;
    if (host === "outlook") {
        cleanJson = cleanUpOutlookJson(json);
    } else {
        cleanJson = cleanUpRichApiJson(json);
    }

    fsx.writeFileSync(`${jsonPath}/${fileName}`, cleanJson);
    console.log(`\nCompleted ${host}`);
    let currentRelease;
    if (host === "outlook") {
        currentRelease = CURRENT_OUTLOOK_RELEASE;
    } else if (host === "excel") {
        currentRelease = CURRENT_EXCEL_RELEASE;
        // Handle ExcelApiOnline corner case.
        console.log(`\nStarting ${host}_online...`);
        json = fsx.readFileSync(`${jsonPath}_online/${fileName}`).toString();
        fsx.writeFileSync(`${jsonPath}_online/${fileName}`, cleanUpRichApiJson(json));
        console.log(`\nCompleted ${host}_online`);
    } else if (host === "word") {
        currentRelease = CURRENT_WORD_RELEASE;
        // Handle WordApiOnline corner case.
        console.log(`\nStarting ${host}_online...`);
        json = fsx.readFileSync(`${jsonPath}_online/${fileName}`).toString();
        fsx.writeFileSync(`${jsonPath}_online/${fileName}`, cleanUpRichApiJson(json));
        console.log(`\nCompleted ${host}_online`);
        // Handle WordApiHiddenDocument 1.4 case.
        console.log(`\nStarting ${host}_1_4_hidden_document...`);
        json = fsx.readFileSync(`${jsonPath}_1_4_hidden_document/${fileName}`).toString();
        fsx.writeFileSync(`${jsonPath}_1_4_hidden_document/${fileName}`, cleanUpRichApiJson(json));
        console.log(`\nCompleted ${host}_1_4_hidden_document`);
        // Handle WordApiHiddenDocument 1.3 case.
        console.log(`\nStarting ${host}_1_3_hidden_document...`);
        json = fsx.readFileSync(`${jsonPath}_1_3_hidden_document/${fileName}`).toString();
        fsx.writeFileSync(`${jsonPath}_1_3_hidden_document/${fileName}`, cleanUpRichApiJson(json));
        console.log(`\nCompleted ${host}_1_3_hidden_document`);
    } else if (host === "powerpoint") {
        currentRelease = CURRENT_POWERPOINT_RELEASE;
    } else {
        currentRelease = 0;
    }

    if (currentRelease > 0) {
        for (let i = currentRelease; i > 0; i--) {
            console.log(`\nStarting ${host}_1_${i}...`);
            json = fsx.readFileSync(`${jsonPath}_1_${i}/${fileName}`).toString();
            if (host === "outlook") {
                cleanJson = cleanUpOutlookJson(json);
            } else {
                cleanJson = cleanUpRichApiJson(json);
            }
            fsx.writeFileSync(`${jsonPath}_1_${i}/${fileName}`, cleanJson);
            console.log(`Completed ${host}_1_${i}`);
        }
    }

    console.log(`\nCompleted ${host} json cross-referencing cleanup`);
}

function cleanUpOutlookJson(jsonString : string) {
    return jsonString.replace(/(\"CommonAPI\.\w+",[\s]+"canonicalReference": ")outlook!/gm, "$1office!")
                     .replace(/("kind": "EnumMember",((?!kind)[\s\S])+"docComment":.*)@remarks\\n/gm, `$1`);
}

function cleanUpRichApiJson(jsonString : string) {
    return jsonString.replace(/(excel|word|visio|onenote|powerpoint)\!OfficeExtension/g, "office!OfficeExtension")
                     .replace(/("kind": "EnumMember",((?!kind)[\s\S])+"docComment":.*)@remarks\\n/gm, `$1`);
}

function cleanUpOutlookMarkdown(markdownString : string) {
    return markdownString.replace(/CommonAPI/gm, "Office");
}

function writeSnippetFileAndClearYamlIfNew(snippetsFilePath: string, snippetsContent: string, keyword: string) {
    const yamlRoot = "../yaml";

    let existingSnippets = "";
    if (fsx.existsSync(snippetsFilePath)) {
        existingSnippets = fsx.readFileSync(snippetsFilePath).toString();
    }

    if (existingSnippets !== snippetsContent) {
        fsx.writeFileSync(snippetsFilePath, snippetsContent);

        fsx.readdirSync(yamlRoot).forEach((yamlFolder) => {
            if (yamlFolder.indexOf(keyword) >= 0) {
                console.log(`Removing ${yamlRoot}/${yamlFolder}`);
                fsx.removeSync(`${yamlRoot}/${yamlFolder}`);
            }
        });
    }
}

async function tryCatch(call: () => Promise<void>) {
    try {
        await call();
    } catch (e) {
        console.error(e);
        process.exit(1);
    }
}
