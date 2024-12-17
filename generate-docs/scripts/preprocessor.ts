#!/usr/bin/env node --harmony

import { fetchAndThrowOnError, DtsBuilder } from './util';
import { promptFromList } from './simple-prompts';
import * as path from "path";
import * as fsx from 'fs-extra';

tryCatch(async () => {
    const args = process.argv.slice(2);
    let sourceChoice;

    // Bypass the prompt - for use with the GitHub Action.
    if (args.length > 0 && args[0] !== null && args[0].trim().length > 0) {
        sourceChoice = args[0].trim();
        console.log(`Bypassing prompt with source choice ${sourceChoice}`);
    } else {
        // ----
        // Display prompts
        // ----
        console.log('\n\n');

        console.log('\n');
        sourceChoice = await promptFromList({
            message: `What is the source of the Office-js TypeScript definition files that should be used to generate the docs?`,
            choices: [
                { name: "DefinitelyTyped (optimized rebuild)", value: "DT" },
                { name: "DefinitelyTyped (full rebuild)", value: "DT+" },
                { name: "CDN (if available)", value: "CDN" },
                { name: "Local files [generate-docs\\script-inputs\\*.d.ts]", value: "Local" }
            ]
        });
    }


    let urlToCopyOfficeJsFrom = "";
    let urlToCopyPreviewOfficeJsFrom = "";
    let urlToCopyCustomFunctionsRuntimeFrom = "";
    let urlToCopyOfficeRuntimeFrom = "";
    let forceRebuild = true;

    switch (sourceChoice) {
        case "DT":
            forceRebuild = false;
        case "DT+":
            urlToCopyOfficeJsFrom = "https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts";
            urlToCopyPreviewOfficeJsFrom = "https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts";
            urlToCopyCustomFunctionsRuntimeFrom = "https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/custom-functions-runtime/index.d.ts";
            urlToCopyOfficeRuntimeFrom = "https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-runtime/index.d.ts";
            break;
        case "CDN":
            urlToCopyOfficeJsFrom = "https://res-sdp.public.cdn.office.net/appsforoffice/_1cdn_bucketedcontent/lib/1.1/hosted/office.d.ts";
            urlToCopyPreviewOfficeJsFrom = "https://res-sdp.public.cdn.office.net/appsforoffice/_1cdn_bucketedcontent/lib/beta/hosted/office.d.ts";
            urlToCopyCustomFunctionsRuntimeFrom = "https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/custom-functions-runtime/index.d.ts";
            urlToCopyOfficeRuntimeFrom = "https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-runtime/index.d.ts";
            break;
        // Note: Using 1CDN instead of "appsforoffice.microsoft.com"
        //     to avoid being redirected to the EDOG environment on corpnet.
        // If we ever want to generate not just public d.ts but also "office-with-first-party.d.ts",
        //     replace the filename.
        case "Local":
            break;
        default:
            throw new Error(`Invalid prompt selection: ${sourceChoice}`);
    }

    console.log("\nStarting preprocessor script...\n");

    // ----
    // Process office.d.ts
    // ----
    const localReleaseDtsPath = "../script-inputs/office.d.ts";
    if (urlToCopyOfficeJsFrom.length > 0) {
        console.log(`Pulling Office.js TypeScript definition file from: ${urlToCopyOfficeJsFrom}`);
        fsx.writeFileSync(localReleaseDtsPath, await fetchAndThrowOnError(urlToCopyOfficeJsFrom, "text"));
    }

    const localPreviewDtsPath = "../script-inputs/office_preview.d.ts";
    if (urlToCopyPreviewOfficeJsFrom.length > 0) {
        console.log(`Pulling Office.js (preview) TypeScript definition file from: ${urlToCopyPreviewOfficeJsFrom}`);
        fsx.writeFileSync(localPreviewDtsPath, await fetchAndThrowOnError(urlToCopyPreviewOfficeJsFrom, "text"));
    }

    let releaseDefinitions = cleanUpDts(localReleaseDtsPath);
    let previewDefinitions = cleanUpDts(localPreviewDtsPath);

    const dtsBuilder = new DtsBuilder();

    console.log("\nCreating separate d.ts files...");

    console.log("create file: office.d.ts (preview)");
    makeDtsAndClearJsonIfNew(
        '../api-extractor-inputs-office/office.d.ts',
        handleCommonImports(dtsBuilder.extractDtsSection(previewDefinitions, "Begin Office namespace", "End Office namespace") +
            '\n' +
            '\n' +
            dtsBuilder.extractDtsSection(releaseDefinitions, "Begin OfficeExtension runtime", "End OfficeExtension runtime"), "Common API"),
        "office",
        forceRebuild
    );

    console.log("create file: office.d.ts");
    makeDtsAndClearJsonIfNew(
        '../api-extractor-inputs-office-release/office.d.ts',
        handleCommonImports(dtsBuilder.extractDtsSection(releaseDefinitions, "Begin Office namespace", "End Office namespace") +
            '\n' +
            '\n' +
            dtsBuilder.extractDtsSection(releaseDefinitions, "Begin OfficeExtension runtime", "End OfficeExtension runtime"), "Common API"),
        "office",
        forceRebuild
    );

    console.log("\ncreate file: excel.d.ts (preview)");
    makeDtsAndClearJsonIfNew(
        '../api-extractor-inputs-excel/excel.d.ts',
        handleCommonImports(handleLiteralParameterOverloads(excelSpecificCleanup(dtsBuilder.extractDtsSection(previewDefinitions, "Begin Excel APIs", "End Excel APIs"))), "Other"),
        "excel",
        forceRebuild
    );

    console.log("\ncreate file: excel.d.ts (release)");
    makeDtsAndClearJsonIfNew(
        '../api-extractor-inputs-excel-release/excel_online/excel.d.ts',
        handleCommonImports(handleLiteralParameterOverloads(excelSpecificCleanup(dtsBuilder.extractDtsSection(releaseDefinitions, "Begin Excel APIs", "End Excel APIs"))), "Other", true),
        "excel",
        forceRebuild
    );

    console.log("create file: onenote.d.ts");
    makeDtsAndClearJsonIfNew(
        '../api-extractor-inputs-onenote/onenote.d.ts',
        handleCommonImports(handleLiteralParameterOverloads(dtsBuilder.extractDtsSection(releaseDefinitions, "Begin OneNote APIs", "End OneNote APIs")), "Other"),
        "onenote",
        forceRebuild
    );

    console.log("create file: outlook.d.ts (preview)");
    makeDtsAndClearJsonIfNew(
        '../api-extractor-inputs-outlook/outlook.d.ts',
        handleCommonImports(dtsBuilder.extractDtsSection(previewDefinitions, "Begin Exchange APIs", "End Exchange APIs"), "Outlook"),
        "outlook",
        forceRebuild
    );

    console.log("create file: outlook.d.ts (release)");
    makeDtsAndClearJsonIfNew(
        '../api-extractor-inputs-outlook-release/outlook_1_14/outlook.d.ts',
        handleCommonImports(dtsBuilder.extractDtsSection(releaseDefinitions, "Begin Exchange APIs", "End Exchange APIs"), "Outlook", true),
        "outlook",
        forceRebuild
    );

    console.log("create file: powerpoint.d.ts (preview)");
    makeDtsAndClearJsonIfNew(
        '../api-extractor-inputs-powerpoint/powerpoint.d.ts',
        handleCommonImports(handleLiteralParameterOverloads(dtsBuilder.extractDtsSection(previewDefinitions, "Begin PowerPoint APIs", "End PowerPoint APIs")), "Other"),
        "powerpoint",
        forceRebuild
    );

    console.log("create file: powerpoint.d.ts (release)");
    makeDtsAndClearJsonIfNew(
        '../api-extractor-inputs-powerpoint-release/PowerPoint_1_7/powerpoint.d.ts',
        handleCommonImports(handleLiteralParameterOverloads(dtsBuilder.extractDtsSection(releaseDefinitions, "Begin PowerPoint APIs", "End PowerPoint APIs")), "Other", true),
        "powerpoint",
        forceRebuild
    );

    console.log("create file: visio.d.ts");
    makeDtsAndClearJsonIfNew(
        '../api-extractor-inputs-visio/visio.d.ts',
        handleCommonImports(handleLiteralParameterOverloads(dtsBuilder.extractDtsSection(releaseDefinitions, "Begin Visio APIs", "End Visio APIs")), "Other"),
        "visio",
        forceRebuild
    );

    console.log("create file: word.d.ts (preview)");
    makeDtsAndClearJsonIfNew(
        '../api-extractor-inputs-word/word.d.ts',
        handleCommonImports(handleLiteralParameterOverloads(wordSpecificCleanup(dtsBuilder.extractDtsSection(previewDefinitions, "Begin Word APIs", "End Word APIs"))), "Other"),
        "word",
        forceRebuild
    );

    console.log("\ncreate file: word.d.ts (release)");
    makeDtsAndClearJsonIfNew(
        '../api-extractor-inputs-word-release/word_online/word-init.d.ts',
        handleCommonImports(handleLiteralParameterOverloads(wordSpecificCleanup(dtsBuilder.extractDtsSection(releaseDefinitions, "Begin Word APIs", "End Word APIs"))), "Other", true),
        "word",
        forceRebuild
    );

    // ----
    // Process Custom Functions d.ts
    // ----
    if (urlToCopyCustomFunctionsRuntimeFrom.length > 0) {
        console.log(`Pulling Custom Functions TypeScript definition file from: ${urlToCopyCustomFunctionsRuntimeFrom}`);
        fsx.writeFileSync("../script-inputs/custom-functions-runtime.d.ts", await fetchAndThrowOnError(urlToCopyCustomFunctionsRuntimeFrom, "text"));
    }
    console.log(`\nReading from ${path.resolve("../script-inputs/custom-functions-runtime.d.ts")}`);
    let definitionsForCfs : string = fsx.readFileSync("../script-inputs/custom-functions-runtime.d.ts").toString();

    console.log("Fixing issues with d.ts file...");
    definitionsForCfs = applyRegularExpressions(definitionsForCfs);

    console.log("create file: custom-functions-runtime.d.ts");
    makeDtsAndClearJsonIfNew('../api-extractor-inputs-custom-functions-runtime/custom-functions-runtime.d.ts', definitionsForCfs, "excel", forceRebuild);

    // ----
    // Process Office Runtime d.ts
    // ----
    if (urlToCopyOfficeRuntimeFrom.length > 0) {
        console.log(`Pulling Office Runtime TypeScript definition file from: ${urlToCopyOfficeRuntimeFrom}`);
        fsx.writeFileSync("../script-inputs/office-runtime.d.ts", await fetchAndThrowOnError(urlToCopyOfficeRuntimeFrom, "text"));
    }
    console.log(`\nReading from ${path.resolve("../script-inputs/office-runtime.d.ts")}`);
    let definitionsForORun : string = fsx.readFileSync("../script-inputs/office-runtime.d.ts").toString();

    console.log("Fixing issues with d.ts file...");
    definitionsForORun = applyRegularExpressions(definitionsForORun);

    console.log("create file: office-runtime.d.ts");
    makeDtsAndClearJsonIfNew('../api-extractor-inputs-office-runtime/office-runtime.d.ts', definitionsForORun, "office", forceRebuild);

    console.log("\nPreprocessor script complete!");

    process.exit(0);
});

function excelSpecificCleanup(dtsContent: string) {
    return dtsContent.replace(/export interface .*Set {\r?\n.*Icon;/gm, `/** [Api set: ExcelApi 1.2] */\n\t$&`)
        .replace("export interface IconCollections {", "/** [Api set: ExcelApi 1.2] */\n\texport interface IconCollections {")
        .replace("var icons: IconCollections;", "/** [Api set: ExcelApi 1.2] */\n\tvar icons: IconCollections;");
}

function wordSpecificCleanup(dtsContent: string) {
    return dtsContent.replace("readonly application: Application;", "/** [Api set: WordApi 1.3] **/\n\t\treadonly application: Application;");
}

function cleanUpDts(localDtsPath: string): string {
    console.log(`\nReading from ${path.resolve(localDtsPath)}`);
    let definitions = fsx.readFileSync(localDtsPath).toString();

    console.log("\nFixing issues with d.ts file...");
    return applyRegularExpressions(definitions.replace(/(extends OfficeCore.RequestContext)/g, `extends OfficeExtension.ClientRequestContext`));
}


// ----
// Helper function to apply regular expressions to d.ts file contents
// ----
function applyRegularExpressions (definitionsIn) {
    return definitionsIn.replace(/^(\s*)(declare namespace)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(declare module)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(namespace)(\s+)(?!=)/gm, `$1export $2$3`)
        .replace(/^(\s*)(class)(\s+)(?!=)/gm, `$1export $2$3`)
        .replace(/^(\s*)(interface)(\s+)(?!=)/gm, `$1export $2$3`)
        .replace(/^(\s*)(module)(\s+)(?!=)/gm, `$1export $2$3`)
        .replace(/^(\s*)(function)(\s+)(?!=)/gm, `$1export $2$3`)
        .replace(/^(\s*)(type)(\s+)(?!=)/gm, `$1export $2$3`)
        .replace(/(\s*)(@param)(\s+)(\w+)(\s)(\s)/g, `$1$2$3$4$5`)
        .replace(/(\s*)(@param)(\s+)(\w+)(\s+)([^\-])/g, `$1$2$3$4$5- $6`);
}

function handleCommonImports(hostDts: string, hostName: "Common API" | "Outlook" | "Other", isVersioned?: boolean): string {
    const commonApiNamespaceImport = "import \{ OfficeExtension \} from \"" + (isVersioned ? "../" : "") + "../api-extractor-inputs-office/office\"\n";
    const outlookApiNamespaceImport = "import \{ Office as Outlook\} from \"" + (isVersioned ? "../" : "") + "../api-extractor-inputs-outlook/outlook\"\n";
    const commonApiNamespaceImportForOutlook = "import \{Office as CommonAPI\} from \"" + (isVersioned ? "../" : "") + "../api-extractor-inputs-office/office\"\n";
    if (hostName === "Outlook") {
        hostDts = hostDts.replace(/: Office\./g, ": CommonAPI.")
                         .replace(/\<Office\./g, "<CommonAPI.");
        return commonApiNamespaceImportForOutlook + hostDts;
    } else if (hostName === "Common API") {
        hostDts = hostDts.replace(/Office\.Mailbox/g, "Outlook.Mailbox")
                         .replace(/Office\.RoamingSettings/g, "Outlook.RoamingSettings")
                         .replace(/Office\.SensitivityLabelsCatalog/g, "Outlook.SensitivityLabelsCatalog");
        return outlookApiNamespaceImport + hostDts;
    } else {
        hostDts = hostDts.replace(/Office\.Mailbox/g, "Outlook.Mailbox")
                         .replace(/Office\.RoamingSettings/g, "Outlook.RoamingSettings")
                         .replace(/Office\.SensitivityLabelsCatalog/g, "Outlook.SensitivityLabelsCatalog");
        return commonApiNamespaceImport + outlookApiNamespaceImport + hostDts;
    }
}

function handleLiteralParameterOverloads(dtsString: string): string {
    // rename parameters for string literal overloads
    const matches = dtsString.match(/([a-zA-Z]+)(\??:)([\n]?([ |]*\"[\w]*\"[|,\n]*)+?)([ ]*[\),])/g);
    let matchIndex = 0;
    if (matches) {
        matches.forEach((match) => {
            let parameterName = match.substring(0, match.indexOf(":"));
            matchIndex = dtsString.indexOf(match, matchIndex);
            parameterName = parameterName.indexOf("?") >= 0 ? parameterName.substring(0, parameterName.length - 1) : parameterName;
            const parameterString = `@param ${parameterName} `;
            const index = dtsString.lastIndexOf(parameterString, matchIndex);
            if (index < 0) {
                // Only warn if this wasn't found in a comment.
                if (dtsString.substring(dtsString.lastIndexOf("*", matchIndex), matchIndex + match.length).match(/([\*]+)/g).length == 0) {
                    console.warn("Missing @param for literal parameter: " + match);
                }
            } else {
                dtsString = dtsString.substring(0, index)
                + "@param " + parameterName + "String "
                + dtsString.substring(index + parameterString.length);
                matchIndex += match.length;
            }
        });
    }
    return dtsString.replace(/([a-zA-Z]+)(\??:)([\n]?([ |]*\"[\w]*\"[|,\n]*)+?)([ ]*[\),])/g, "$1String$2$3$5").replace(/([\*]+.+ .+)String:/g, "$1:");
}

function makeDtsAndClearJsonIfNew(dtsFilePath: string, dtsContent: string, keyword: string, forceNew: boolean) {
    const jsonRoot = "../json";
    const yamlRoot = "../yaml";
    
    let existingDts = fsx.readFileSync(dtsFilePath).toString();
    if (existingDts !== dtsContent || forceNew) {
        fsx.writeFileSync(dtsFilePath, dtsContent);
        
        fsx.readdirSync(jsonRoot).forEach((jsonFolder) => {
            if (jsonFolder.indexOf(keyword) >= 0) {
                console.log(`Removing ${jsonRoot}/${jsonFolder}`);
                fsx.removeSync(`${jsonRoot}/${jsonFolder}`);
            }
        });
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