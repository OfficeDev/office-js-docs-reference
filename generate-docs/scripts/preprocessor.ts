#!/usr/bin/env node --harmony

import { fetchAndThrowOnError, DtsBuilder } from './util';
import { promptFromList } from './simple-prompts';
import * as path from "path";
import * as fsx from 'fs-extra';

tryCatch(async () => {
    // ----
    // Display prompts
    // ----
    console.log('\n\n');

    console.log('\n');
    const sourceChoice = await promptFromList({
        message: `What is the source of the Office-js TypeScript definition files that should be used to generate the docs?`,
        choices: [
            { name: "DefinitelyTyped", value: "DT" },
            { name: "CDN (if available)", value: "CDN" },
            { name: "Local files [generate-docs\\script-inputs\\*.d.ts]", value: "Local" }
        ]
    });


    let urlToCopyOfficeJsFrom = "";
    let urlToCopyPreviewOfficeJsFrom = "";
    let urlToCopyCustomFunctionsRuntimeFrom = "";
    let urlToCopyOfficeRuntimeFrom = "";

    switch (sourceChoice) {
        case "DT":
            urlToCopyOfficeJsFrom = "https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts";
            urlToCopyPreviewOfficeJsFrom = "https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts";
            urlToCopyCustomFunctionsRuntimeFrom = "https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/custom-functions-runtime/index.d.ts";
            urlToCopyOfficeRuntimeFrom = "https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-runtime/index.d.ts";
            break;
        case "CDN":
            urlToCopyOfficeJsFrom = "https://appsforoffice.officeapps.live.com/lib/1.1/hosted/office.d.ts";
            urlToCopyPreviewOfficeJsFrom = "https://appsforoffice.officeapps.live.com/lib/beta/hosted/office.d.ts";
            urlToCopyCustomFunctionsRuntimeFrom = "https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/custom-functions-runtime/index.d.ts";
            urlToCopyOfficeRuntimeFrom = "https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-runtime/index.d.ts";
            break;
        // Note: using "appsforoffice.officeapps.live.com" instead of "appsforoffice.microsoft.com"
        //     to avoid being redirected to the EDOG environment on corpnet.
        // If we ever want to generate not just public d.ts but also "office-with-first-party.d.ts",
        //     replace the filename.
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

    console.log("create file: office.d.ts");
    makeDtsAndClearJsonIfNew(
        '../api-extractor-inputs-office/office.d.ts',
        handleCommonImports(dtsBuilder.extractDtsSection(previewDefinitions, "Begin Office namespace", "End Office namespace") +
            '\n' +
            '\n' +
            dtsBuilder.extractDtsSection(releaseDefinitions, "Begin OfficeExtension runtime", "End OfficeExtension runtime"), "Common API"),
        "office"
    );

    console.log("\ncreate file: excel.d.ts (preview)");
    makeDtsAndClearJsonIfNew(
        '../api-extractor-inputs-excel/excel.d.ts',
        handleCommonImports(handleLiteralParameterOverloads(excelSpecificCleanup(dtsBuilder.extractDtsSection(previewDefinitions, "Begin Excel APIs", "End Excel APIs"))), "Other"),
        "excel"
    );

    console.log("\ncreate file: excel.d.ts (release)");
    makeDtsAndClearJsonIfNew(
        '../api-extractor-inputs-excel-release/excel_online/excel.d.ts',
        handleCommonImports(handleLiteralParameterOverloads(excelSpecificCleanup(dtsBuilder.extractDtsSection(releaseDefinitions, "Begin Excel APIs", "End Excel APIs"))), "Other", true),
        "excel"
    );

    console.log("create file: onenote.d.ts");
    makeDtsAndClearJsonIfNew(
        '../api-extractor-inputs-onenote/onenote.d.ts',
        handleCommonImports(handleLiteralParameterOverloads(dtsBuilder.extractDtsSection(releaseDefinitions, "Begin OneNote APIs", "End OneNote APIs")), "Other"),
        "onenote"
    );

    console.log("create file: outlook.d.ts (preview)");
    makeDtsAndClearJsonIfNew(
        '../api-extractor-inputs-outlook/outlook.d.ts',
        handleCommonImports(dtsBuilder.extractDtsSection(previewDefinitions, "Begin Exchange APIs", "End Exchange APIs"), "Outlook"),
        "outlook"
    );

    console.log("create file: outlook.d.ts (release)");
    makeDtsAndClearJsonIfNew(
        '../api-extractor-inputs-outlook-release/outlook_1_10/outlook.d.ts',
        handleCommonImports(dtsBuilder.extractDtsSection(releaseDefinitions, "Begin Exchange APIs", "End Exchange APIs"), "Outlook", true),
        "outlook"
    );

    console.log("create file: powerpoint.d.ts (preview)");
    makeDtsAndClearJsonIfNew(
        '../api-extractor-inputs-powerpoint/powerpoint.d.ts',
        handleCommonImports(handleLiteralParameterOverloads(dtsBuilder.extractDtsSection(previewDefinitions, "Begin PowerPoint APIs", "End PowerPoint APIs")), "Other"),
        "powerpoint"
    );

    console.log("create file: powerpoint.d.ts (release)");
    makeDtsAndClearJsonIfNew(
        '../api-extractor-inputs-powerpoint-release/PowerPoint_1_2/powerpoint.d.ts',
        handleCommonImports(handleLiteralParameterOverloads(dtsBuilder.extractDtsSection(releaseDefinitions, "Begin PowerPoint APIs", "End PowerPoint APIs")), "Other", true),
        "powerpoint"
    );

    console.log("create file: visio.d.ts");
    makeDtsAndClearJsonIfNew(
        '../api-extractor-inputs-visio/visio.d.ts',
        handleCommonImports(handleLiteralParameterOverloads(dtsBuilder.extractDtsSection(releaseDefinitions, "Begin Visio APIs", "End Visio APIs")), "Other"),
        "visio"
    );

    console.log("create file: word.d.ts (preview)");
    makeDtsAndClearJsonIfNew(
        '../api-extractor-inputs-word/word.d.ts',
        handleCommonImports(handleLiteralParameterOverloads(wordSpecificCleanup(dtsBuilder.extractDtsSection(previewDefinitions, "Begin Word APIs", "End Word APIs"))), "Other"),
        "word"
    );

    console.log("\ncreate file: word.d.ts (release)");
    makeDtsAndClearJsonIfNew(
        '../api-extractor-inputs-word-release/word_1_3/word.d.ts',
        handleCommonImports(handleLiteralParameterOverloads(wordSpecificCleanup(dtsBuilder.extractDtsSection(releaseDefinitions, "Begin Word APIs", "End Word APIs"))), "Other", true),
        "word"
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
    makeDtsAndClearJsonIfNew('../api-extractor-inputs-custom-functions-runtime/custom-functions-runtime.d.ts', definitionsForCfs, "custom");

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
    makeDtsAndClearJsonIfNew('../api-extractor-inputs-office-runtime/office-runtime.d.ts', definitionsForORun, "runtime");

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
    return applyRegularExpressions(
        definitions
        .replace(/([ ]*)load\(option\?: string \| string\[\]\): (Excel|Word|OneNote|Visio|PowerPoint)\.(.*);/g,
                 "$1/**\n$1 * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.\n$1 * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.\n$1 */\n$1load(propertyNames?: string | string[]): $2.$3;")
        .replace(/([ ]*)load\(option\?: {\n[ ]*select\?: string;\n[ ]*expand\?: string;\n[ ]*}\): (Excel|Word|OneNote|Visio|PowerPoint)\.(.*);/gm,
                 "$1/**\n$1 * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.\n$1 * @param propertyNamesAndPaths - Where propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.\n$1 */\n$1load(propertyNamesAndPaths?: { select?: string; expand?: string; }): $2.$3;")
        .replace(/([ ]*)load\(option\?: (Excel|Word|OneNote|Visio|PowerPoint)\.Interfaces\.(.*)CollectionLoadOptions & [Excel|Word|OneNote|Visio|PowerPoint]\.Interfaces\.CollectionLoadOptions\): [Excel|Word|OneNote|Visio|PowerPoint]\.[.*]Collection;/g,
                 "$1/**\n$1 * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.\n$1 * @param collectionLoadOptions - Where collectionLoadOptions.select is a comma-delimited string that specifies the properties to load, and collectionLoadOptions.expand is a comma-delimited string that specifies the navigation properties to load. collectionLoadOptions.top specifies the maximum number of collection items that can be included in the result. collectionLoadOptions.skip specifies the number of items that are to be skipped and not included in the result. If collectionLoadOptions.top is specified, the result set will start after skipping the specified number of items.\n$1 */\n$1load(collectionLoadOptions?: $2.Interfaces.$3CollectionLoadOptions & $2.Interfaces.CollectionLoadOptions): $2.$3Collection;")
        .replace(/(extends OfficeCore.RequestContext)/g, `extends OfficeExtension.ClientRequestContext`)
        .replace(/OfficeExtension\.IPromise\<T\>/g, `Promise<T>`) /* Not needed once type support is added to API Documenter and the OPS YAML schema. */);
}


// ----
// Helper function to apply regular expressions to d.ts file contents
// ----
function applyRegularExpressions (definitionsIn) {
    return definitionsIn.replace(/^(\s*)(declare namespace)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(declare module)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(namespace)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(class)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(interface)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(module)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(function)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(type)(\s+)/gm, `$1export $2$3`)
        .replace(/(\s*)(@param)(\s+)(\w+)(\s)(\s)/g, `$1$2$3$4$5`)
        .replace(/(\s*)(@param)(\s+)(\w+)(\s+)([^\-])/g, `$1$2$3$4$5- $6`);
}

function handleCommonImports(hostDts: string, hostName: "Common API" | "Outlook" | "Other", isVersioned?: boolean): string {
    const commonApiNamespaceImport = "import \{ OfficeExtension \} from \"" + (isVersioned ? "../" : "") + "../api-extractor-inputs-office/office\"\n";
    const outlookApiNamespaceImport = "import \{ Office as Outlook\} from \"" + (isVersioned ? "../" : "") + "../api-extractor-inputs-outlook/outlook\"\n";
    const commonApiNamespaceImportForOutlook = "import \{Office as CommonAPI\} from \"" + (isVersioned ? "../" : "") + "../api-extractor-inputs-office/office\"\n";
    if (hostName === "Outlook") {
        hostDts = hostDts.replace(/: Office\./g, ": CommonAPI.").replace(/\<Office\./g, "<CommonAPI.");
        return commonApiNamespaceImportForOutlook + hostDts;
    } else if (hostName === "Common API") {
        hostDts = hostDts.replace(/Office\.Mailbox/g, "Outlook.Mailbox").replace(/Office\.RoamingSettings/g, "Outlook.RoamingSettings");
        return outlookApiNamespaceImport + hostDts;
    } else {
        hostDts = hostDts.replace(/Office\.Mailbox/g, "Outlook.Mailbox").replace(/Office\.RoamingSettings/g, "Outlook.RoamingSettings");
        return commonApiNamespaceImport + outlookApiNamespaceImport + hostDts;
    }
}

function handleLiteralParameterOverloads(dtsString: string): string {
    // rename parameters for string literal overloads
    const matches = dtsString.match(/([a-zA-Z]+)\??: (\"[a-zA-Z]*\").*:/g);
    let matchIndex = 0;
    if (matches) {
        matches.forEach((match) => {
            let parameterName = match.substring(0, match.indexOf(": "));
            matchIndex = dtsString.indexOf(match, matchIndex);
            parameterName = parameterName.indexOf("?") >= 0 ? parameterName.substring(0, parameterName.length - 1) : parameterName;
            const parameterString = "@param " + parameterName + " ";
            const index = dtsString.lastIndexOf(parameterString, matchIndex);
            if (index < 0) {
                console.warn("Missing @param for literal parameter: " + match);
            } else {
            dtsString = dtsString.substring(0, index)
            + "@param " + parameterName + "String "
            + dtsString.substring(index + parameterString.length);
            matchIndex += match.length;
            }
        });
    }

    return dtsString.replace(/([a-zA-Z]+)(\??: \"[a-zA-Z]*\".*:)/g, "$1String$2");
}

function makeDtsAndClearJsonIfNew(dtsFilePath: string, dtsContent: string, keyword: string) {
    const jsonRoot = "../json";
    const yamlRoot = "../yaml";
    
    let existingDts = fsx.readFileSync(dtsFilePath).toString();
    if (existingDts !== dtsContent) {
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