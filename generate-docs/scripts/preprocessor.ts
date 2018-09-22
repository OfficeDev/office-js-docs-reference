#!/usr/bin/env node --harmony

import { fetchAndThrowOnError, DtsBuilder } from './util';
import { promptFromList } from './simple-prompts';
import * as path from "path";
import * as fsx from 'fs-extra';
import yaml = require('js-yaml');

tryCatch(async () => {
    console.log('\n\n');
    const urlToCopyOfficeJsFrom = await promptFromList({
        message: `What is the source of the office.d.ts file that should be used to generate the docs?`,
        choices: [
            { name: "DefinitelyTyped", value: "https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts" },
            { name: "Prod CDN", value: "https://appsforoffice.officeapps.live.com/lib/1.1/hosted/office.d.ts" },
            { name: "Beta CDN", value: "https://appsforoffice.officeapps.live.com/lib/beta/hosted/office.d.ts" },
            { name: "Local file [generate-docs\\script-inputs\\office.d.ts]", value: "" }
        ]

        // Note: using "appsforoffice.officeapps.live.com" instead of "appsforoffice.microsoft.com"
        //     to avoid being redirected to the EDOG environment on corpnet.
        // If we ever want to generate not just public d.ts but also "office-with-first-party.d.ts",
        //     replace the filename.
    });

    console.log('\n');
    const includeScriptLabSnippets = await promptFromList({
        message: `Do you want to include Script Lab code snippets in the generated docs?`,
        choices: [
            { name: "Yes", value: "y" },
            { name: "No", value: "n" }
        ]
    });

    console.log("\nStarting preprocessor script...");

    // process office.d.ts
    if (urlToCopyOfficeJsFrom.length > 0) {
        fsx.writeFileSync("../script-inputs/office.d.ts", await fetchAndThrowOnError(urlToCopyOfficeJsFrom, "text"));
    }
    console.log(`\nReading from ${path.resolve("../script-inputs/office.d.ts")}`);
    let definitions = fsx.readFileSync("../script-inputs/office.d.ts").toString();

    console.log("\nFixing issues with d.ts file...");
    // Note: This step fixing formatting discrepancies and hiding content we do not wish to expose.
    // LoadOptions are removed, and the corresponding comments are modified to reflect a different overload.
    // set() is removed from RichAPI, along with corresponding comments. This is to reduce traffic to the method pending a decision about modifying the underlying behavior.
    definitions = definitions.replace(/^(\s*)(declare namespace)(\s+)/gm, `$1export $2$3`)
        .replace(/\s*\* \@param options Provides options for which properties of the object to load\.(\s*\*\/)/gm, ' @param option A comma-delimited string or an array of strings that specify the properties to load.$1\n')
        .replace(/\s*load\(option\?: (Excel|Word|OneNote|Visio)\.Interfaces\.\S*LoadOptions.*\): \S*?;/gm, '')
        .replace(/\*\s*?`load\(option\?: string \| string\[\]\): (Excel|Word|OneNote|Visio)\..*?` - Where option is a comma-delimited string or an array of strings that specify the properties to load\./g, '')
        .replace(/interface .*?LoadOptions \{[^}]*?}/gm, '')
        .replace(/\/\*\* Sets multiple properties.*\s*\*\s*\*.@remarks\s*\*\s*\* This method has the following additional signature:\s*\*\s*\* \`set\(properties:.*\s*\*\s*\* @param.*\s*\*.*\s*\*\/\s*set\(properties:.*\s*\/\*\* Sets multiple properties.*\s*set\(properties:.*;/gm, '')
        .replace(/^(\s*)(declare module)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(namespace)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(class)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(interface)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(module)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(function)(\s+)/gm, `$1export $2$3`)
        .replace(/(extends OfficeCore.RequestContext)/g, `extends OfficeExtension.ClientRequestContext`)
        .replace(/(\s*)(@param)(\s+)(\w+)(\s)(\s)/g, `$1$2$3$4$5`)
        .replace(/(\s*)(@param)(\s+)(\w+)(\s+)([^\-])/g, `$1$2$3$4$5- $6`);

    const dtsBuilder = new DtsBuilder();

    console.log("\nCreating separate d.ts files...");

    console.log("\ncreate file: excel.d.ts");
    fsx.writeFileSync(
        '../api-extractor-inputs-excel/excel.d.ts',
        dtsBuilder.extractDtsSection(definitions, "Begin Excel APIs", "End Excel APIs")
    );

    console.log("create file: office.d.ts");
    fsx.writeFileSync(
        '../api-extractor-inputs-office/office.d.ts',
        dtsBuilder.extractDtsSection(definitions, "Begin Office namespace", "End Office namespace") +
        '\n' +
        '\n' +
        dtsBuilder.extractDtsSection(definitions, "Begin OfficeExtension runtime", "End OfficeExtension runtime")
    );

    console.log("create file: onenote.d.ts");
    fsx.writeFileSync(
        '../api-extractor-inputs-onenote/onenote.d.ts',
        dtsBuilder.extractDtsSection(definitions, "Begin OneNote APIs", "End OneNote APIs")
    );

    console.log("create file: outlook.d.ts");
    fsx.writeFileSync(
        '../api-extractor-inputs-outlook/outlook.d.ts',
        dtsBuilder.extractDtsSection(definitions, "Begin Exchange APIs", "End Exchange APIs")
    );

    console.log("create file: visio.d.ts");
    fsx.writeFileSync(
        '../api-extractor-inputs-visio/visio.d.ts',
        dtsBuilder.extractDtsSection(definitions, "Begin Visio APIs", "End Visio APIs")
    );

    console.log("create file: word.d.ts");
    fsx.writeFileSync(
        '../api-extractor-inputs-word/word.d.ts',
        dtsBuilder.extractDtsSection(definitions, "Begin Word APIs", "End Word APIs")
    );

    // process Custom Functions d.ts
    const dtsForCfs = "https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/custom-functions-runtime/index.d.ts";
    fsx.writeFileSync("../script-inputs/custom-functions-runtime.d.ts", await fetchAndThrowOnError(dtsForCfs, "text"));
    console.log(`\nReading from ${path.resolve("../script-inputs/custom-functions-runtime.d.ts")}`);
    let definitionsForCfs : string = fsx.readFileSync("../script-inputs/custom-functions-runtime.d.ts").toString();

    console.log("\nFixing issues with d.ts file...");
    definitionsForCfs = definitionsForCfs.replace(/^(\s*)(declare namespace)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(declare module)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(namespace)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(class)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(interface)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(module)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(function)(\s+)/gm, `$1export $2$3`)
        .replace(/(\s*)(@param)(\s+)(\w+)(\s)(\s)/g, `$1$2$3$4$5`)
        .replace(/(\s*)(@param)(\s+)(\w+)(\s+)([^\-])/g, `$1$2$3$4$5- $6`);

    console.log("create file: custom-functions-runtime.d.ts");
    fsx.writeFileSync('../api-extractor-inputs-custom-functions-runtime/custom-functions-runtime.d.ts', definitionsForCfs);

    // process Office Runtime d.ts
    const dtsForORun = "https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-runtime/index.d.ts";
    fsx.writeFileSync("../script-inputs/office-runtime.d.ts", await fetchAndThrowOnError(dtsForORun, "text"));
    console.log(`\nReading from ${path.resolve("../script-inputs/office-runtime.d.ts")}`);
    let definitionsForORun : string = fsx.readFileSync("../script-inputs/office-runtime.d.ts").toString();

    console.log("\nFixing issues with d.ts file...");
    definitionsForORun = definitionsForORun.replace(/^(\s*)(declare namespace)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(declare module)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(namespace)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(class)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(interface)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(module)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(function)(\s+)/gm, `$1export $2$3`)
        .replace(/(\s*)(@param)(\s+)(\w+)(\s)(\s)/g, `$1$2$3$4$5`)
        .replace(/(\s*)(@param)(\s+)(\w+)(\s+)([^\-])/g, `$1$2$3$4$5- $6`);

    console.log("create file: office-runtime.d.ts");
    fsx.writeFileSync('../api-extractor-inputs-office-runtime/office-runtime.d.ts', definitionsForORun);

    console.log("\nRemoving old snippets input files...");

    const scriptInputsPath = path.resolve("../script-inputs");
    fsx.readdirSync(scriptInputsPath)
        .filter(filename => filename.indexOf("snippets") > 0)
        .forEach(filename => fsx.removeSync(scriptInputsPath + '/' + filename));

    console.log("\nCreating snippets file...");

    if (includeScriptLabSnippets === "y") {
        console.log("\nReading from: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/master/snippet-extractor-output/snippets.yaml");
        fsx.writeFileSync("../script-inputs/script-lab-snippets.yaml", await fetchAndThrowOnError("https://raw.githubusercontent.com/OfficeDev/office-js-snippets/master/snippet-extractor-output/snippets.yaml", "text"));
    }

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
    let snippets: Object = yaml.load(localCodeSnippetsString);

    // If including Script Lab snippets, add them to the set. If a duplicate key exists, merge the Script Lab example(s)
    // into the item with the existing key.
    if (includeScriptLabSnippets === "y") {
        let scriptLabSnippets: Object = yaml.load(fsx.readFileSync(`../script-inputs/script-lab-snippets.yaml`).toString());
        for (const key of Object.keys(scriptLabSnippets)) {
            if (snippets[key]) {
                console.log("Combining local and Script Lab snippets for: " + key);
                snippets[key] = snippets[key].concat(scriptLabSnippets[key]);
            } else {
                snippets[key] = scriptLabSnippets[key];
            }
        }
    }

    console.log("\nWriting snippets to: " + path.resolve("../json/snippets.yaml"));

    console.log("\nPreprocessor script complete!");

    process.exit(0);
});


async function tryCatch(call: () => Promise<void>) {
    try {
        await call();
    } catch (e) {
        console.error(e);
        process.exit(1);
    }
}
