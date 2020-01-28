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
    const urlToCopyOfficeJsFrom = await promptFromList({
        message: `What is the source of the Office-js TypeScript definition file that should be used to generate the RELEASE docs?`,
        choices: [
            { name: "DefinitelyTyped", value: "https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts" },
            { name: "Prod CDN", value: "https://appsforoffice.officeapps.live.com/lib/1.1/hosted/office.d.ts" },
            { name: "Local file [generate-docs\\script-inputs\\office.d.ts]", value: "" }
        ]

        // Note: using "appsforoffice.officeapps.live.com" instead of "appsforoffice.microsoft.com"
        //     to avoid being redirected to the EDOG environment on corpnet.
        // If we ever want to generate not just public d.ts but also "office-with-first-party.d.ts",
        //     replace the filename.
    });

    console.log('\n');
    const urlToCopyPreviewOfficeJsFrom = await promptFromList({
        message: `What is the source of the Office-js TypeScript definition file that should be used to generate the PREVIEW docs?`,
        choices: [
            { name: "DefinitelyTyped (preview)", value: "https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts" },
            { name: "Beta CDN", value: "https://appsforoffice.officeapps.live.com/lib/beta/hosted/office.d.ts" },
            { name: "Local file [generate-docs\\script-inputs\\office_preview.d.ts]", value: "" }
        ]
    });

    console.log('\n');
    const urlToCopyCustomFunctionsRuntimeFrom = await promptFromList({
        message: `What is the source of the Custom Functions Runtime TypeScript definition file that should be used to generate the docs?`,
        choices: [
            { name: "DefinitelyTyped", value: "https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/custom-functions-runtime/index.d.ts" },
            { name: "Local file [generate-docs\\script-inputs\\custom-functions-runtime.d.ts]", value: "" }
        ]
    });

    console.log('\n');
    const urlToCopyOfficeRuntimeFrom = await promptFromList({
        message: `What is the source of the Office Runtime TypeScript definition file that should be used to generate the docs?`,
        choices: [
            { name: "DefinitelyTyped", value: "https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-runtime/index.d.ts" },
            { name: "Local file [generate-docs\\script-inputs\\office-runtime.d.ts]", value: "" }
        ]
    });

    console.log("\nStarting preprocessor script...");

    // ----
    // Process office.d.ts
    // ----
    const localReleaseDtsPath = "../script-inputs/office.d.ts";
    if (urlToCopyOfficeJsFrom.length > 0) {
        fsx.writeFileSync(localReleaseDtsPath, await fetchAndThrowOnError(urlToCopyOfficeJsFrom, "text"));
    }

    const localPreviewDtsPath = "../script-inputs/office_preview.d.ts";
    if (urlToCopyPreviewOfficeJsFrom.length > 0) {
        fsx.writeFileSync(localPreviewDtsPath, await fetchAndThrowOnError(urlToCopyPreviewOfficeJsFrom, "text"));
    }

    let releaseDefinitions = cleanUpDts(localReleaseDtsPath);
    let previewDefinitions = cleanUpDts(localPreviewDtsPath);

    const dtsBuilder = new DtsBuilder();

    console.log("\nCreating separate d.ts files...");

    console.log("create file: office.d.ts");
    fsx.writeFileSync(
        '../api-extractor-inputs-office/office.d.ts',
        handleCommonImports(dtsBuilder.extractDtsSection(previewDefinitions, "Begin Office namespace", "End Office namespace") +
        '\n' +
        '\n' +
        dtsBuilder.extractDtsSection(releaseDefinitions, "Begin OfficeExtension runtime", "End OfficeExtension runtime"), "Common API")
    );

    console.log("\ncreate file: excel.d.ts (preview)");
    fsx.writeFileSync(
        '../api-extractor-inputs-excel/excel.d.ts',
        handleCommonImports(handleLiteralParameterOverloads(excelSpecificCleanup(dtsBuilder.extractDtsSection(previewDefinitions, "Begin Excel APIs", "End Excel APIs"))), "Other")
    );

    console.log("\ncreate file: excel.d.ts (release)");
    fsx.writeFileSync(
        '../api-extractor-inputs-excel-release/excel_online/excel.d.ts',
        handleCommonImports(handleLiteralParameterOverloads(excelSpecificCleanup(dtsBuilder.extractDtsSection(releaseDefinitions, "Begin Excel APIs", "End Excel APIs"))), "Other", true)
    );

    console.log("create file: onenote.d.ts");
    fsx.writeFileSync(
        '../api-extractor-inputs-onenote/onenote.d.ts',
        handleCommonImports(handleLiteralParameterOverloads(dtsBuilder.extractDtsSection(releaseDefinitions, "Begin OneNote APIs", "End OneNote APIs")), "Other")
    );

    console.log("create file: outlook.d.ts (preview)");
    fsx.writeFileSync(
        '../api-extractor-inputs-outlook/outlook.d.ts',
        handleCommonImports(outlookSpecificCleanup(dtsBuilder.extractDtsSection(previewDefinitions, "Begin Exchange APIs", "End Exchange APIs")), "Outlook")
    );

    console.log("create file: outlook.d.ts (release)");
    fsx.writeFileSync(
        '../api-extractor-inputs-outlook-release/outlook_1_8/outlook.d.ts',
        handleCommonImports(outlookSpecificCleanup(dtsBuilder.extractDtsSection(releaseDefinitions, "Begin Exchange APIs", "End Exchange APIs")), "Outlook", true)
    );

    console.log("create file: powerpoint.d.ts");
    fsx.writeFileSync(
        '../api-extractor-inputs-powerpoint/powerpoint.d.ts',
        handleCommonImports(dtsBuilder.extractDtsSection(releaseDefinitions, "Begin PowerPoint APIs", "End PowerPoint APIs"), "Other")
    );

    console.log("create file: visio.d.ts");
    fsx.writeFileSync(
        '../api-extractor-inputs-visio/visio.d.ts',
        handleCommonImports(handleLiteralParameterOverloads(dtsBuilder.extractDtsSection(releaseDefinitions, "Begin Visio APIs", "End Visio APIs")), "Other")
    );

    console.log("create file: word.d.ts (preview)");
    fsx.writeFileSync(
        '../api-extractor-inputs-word/word.d.ts',
        handleCommonImports(handleLiteralParameterOverloads(wordSpecificCleanup(dtsBuilder.extractDtsSection(previewDefinitions, "Begin Word APIs", "End Word APIs"))), "Other")
    );

    console.log("\ncreate file: word.d.ts (release)");
    fsx.writeFileSync(
        '../api-extractor-inputs-word-release/word_1_3/word.d.ts',
        handleCommonImports(handleLiteralParameterOverloads(wordSpecificCleanup(dtsBuilder.extractDtsSection(releaseDefinitions, "Begin Word APIs", "End Word APIs"))), "Other", true)
    );

    // ----
    // Process Custom Functions d.ts
    // ----
    if (urlToCopyCustomFunctionsRuntimeFrom.length > 0) {
        fsx.writeFileSync("../script-inputs/custom-functions-runtime.d.ts", await fetchAndThrowOnError(urlToCopyCustomFunctionsRuntimeFrom, "text"));
    }
    console.log(`\nReading from ${path.resolve("../script-inputs/custom-functions-runtime.d.ts")}`);
    let definitionsForCfs : string = fsx.readFileSync("../script-inputs/custom-functions-runtime.d.ts").toString();

    console.log("Fixing issues with d.ts file...");
    definitionsForCfs = applyRegularExpressions(definitionsForCfs);

    console.log("create file: custom-functions-runtime.d.ts");
    fsx.writeFileSync('../api-extractor-inputs-custom-functions-runtime/custom-functions-runtime.d.ts', definitionsForCfs);

    // ----
    // Process Office Runtime d.ts
    // ----
    if (urlToCopyOfficeRuntimeFrom.length > 0) {
        fsx.writeFileSync("../script-inputs/office-runtime.d.ts", await fetchAndThrowOnError(urlToCopyOfficeRuntimeFrom, "text"));
    }
    console.log(`\nReading from ${path.resolve("../script-inputs/office-runtime.d.ts")}`);
    let definitionsForORun : string = fsx.readFileSync("../script-inputs/office-runtime.d.ts").toString();

    console.log("Fixing issues with d.ts file...");
    definitionsForORun = applyRegularExpressions(definitionsForORun);

    console.log("create file: office-runtime.d.ts");
    fsx.writeFileSync('../api-extractor-inputs-office-runtime/office-runtime.d.ts', definitionsForORun);

    console.log("\nPreprocessor script complete!");

    process.exit(0);
});

function excelSpecificCleanup(dtsContent: string) {
    return dtsContent.replace(/export interface .*Set {\r?\n.*Icon;/gm, `/** [Api set: ExcelApi 1.2] */\n\t$&`)
        .replace("export interface IconCollections {", "/** [Api set: ExcelApi 1.2] */\n\texport interface IconCollections {")
        .replace("var icons: IconCollections;", "/** [Api set: ExcelApi 1.2] */\n\tvar icons: IconCollections;");
}

function outlookSpecificCleanup(dtsContent: string) {
    // Use dtsContent to handle Item then merge in other updated objects.
    let dtsContentForAppointment: string = dtsContent;
    let dtsContentForMessage: string = dtsContent;
    let dtsContentForItemCompose: string = dtsContent;
    let dtsContentForItemRead: string = dtsContent;
    let dtsContentForAppointmentCompose: string = dtsContent;
    let dtsContentForAppointmentRead: string = dtsContent;
    let dtsContentForMessageCompose: string = dtsContent;
    let dtsContentForMessageRead: string = dtsContent;

    /** ITEM */
    // Initial Item details.
    let itemInterface = dtsContent.indexOf("interface Item ");
    let itemStart = dtsContent.indexOf("{", itemInterface);
    let itemEnd = dtsContent.indexOf("    }", itemStart);

    // Copy contents of Appointment interface to Item.
    let appointmentInterface = dtsContent.indexOf("interface Appointment ");
    let appointmentStart = dtsContent.indexOf("{", appointmentInterface);
    let appointmentEnd = dtsContent.indexOf("    }", appointmentStart);
    let appointment = dtsContent.substring(appointmentStart + 1, appointmentEnd - 1);

    dtsContent = dtsContent.substring(0, itemEnd - 1) + appointment + dtsContent.substring(itemEnd);

    // Copy contents of Message interface to Item.
    itemEnd = dtsContent.indexOf("    }", itemStart);

    let messageInterface = dtsContent.indexOf("interface Message ");
    let messageStart = dtsContent.indexOf("{", messageInterface);
    let messageEnd = dtsContent.indexOf("    }", messageStart);
    let message = dtsContent.substring(messageStart + 1, messageEnd - 1);

    dtsContent = dtsContent.substring(0, itemEnd - 1) + message + dtsContent.substring(itemEnd);

    // Copy contents of ItemCompose interface to Item.
    itemEnd = dtsContent.indexOf("    }", itemStart);

    let itemComposeInterface = dtsContent.indexOf("interface ItemCompose ");
    let itemComposeStart = dtsContent.indexOf("{", itemComposeInterface);
    let itemComposeEnd = dtsContent.indexOf("    }", itemComposeStart);
    let itemCompose = dtsContent.substring(itemComposeStart + 1, itemComposeEnd - 1);

    dtsContent = dtsContent.substring(0, itemEnd - 1) + itemCompose + dtsContent.substring(itemEnd);

    // Copy contents of ItemRead interface to Item.
    itemEnd = dtsContent.indexOf("    }", itemStart);

    let itemReadInterface = dtsContent.indexOf("interface ItemRead ");
    let itemReadStart = dtsContent.indexOf("{", itemReadInterface);
    let itemReadEnd = dtsContent.indexOf("    }", itemReadStart);
    let itemRead = dtsContent.substring(itemReadStart + 1, itemReadEnd - 1);

    dtsContent = dtsContent.substring(0, itemEnd - 1) + itemRead + dtsContent.substring(itemEnd);

    // Copy contents of AppointmentCompose interface to Item.
    itemEnd = dtsContent.indexOf("    }", itemStart);

    let appointmentComposeInterface = dtsContent.indexOf("interface AppointmentCompose ");
    let appointmentComposeStart = dtsContent.indexOf("{", appointmentComposeInterface);
    let appointmentComposeEnd = dtsContent.indexOf("    }", appointmentComposeStart);
    let appointmentCompose = dtsContent.substring(appointmentComposeStart + 1, appointmentComposeEnd - 1);

    dtsContent = dtsContent.substring(0, itemEnd - 1) + appointmentCompose + dtsContent.substring(itemEnd);

    // Copy contents of AppointmentRead interface to Item.
    itemEnd = dtsContent.indexOf("    }", itemStart);

    let appointmentReadInterface = dtsContent.indexOf("interface AppointmentRead ");
    let appointmentReadStart = dtsContent.indexOf("{", appointmentReadInterface);
    let appointmentReadEnd = dtsContent.indexOf("    }", appointmentReadStart);
    let appointmentRead = dtsContent.substring(appointmentReadStart + 1, appointmentReadEnd - 1);

    dtsContent = dtsContent.substring(0, itemEnd - 1) + appointmentRead + dtsContent.substring(itemEnd);

    // Copy contents of MessageCompose interface to Item.
    itemEnd = dtsContent.indexOf("    }", itemStart);

    let messageComposeInterface = dtsContent.indexOf("interface MessageCompose ");
    let messageComposeStart = dtsContent.indexOf("{", messageComposeInterface);
    let messageComposeEnd = dtsContent.indexOf("    }", messageComposeStart);
    let messageCompose = dtsContent.substring(messageComposeStart + 1, messageComposeEnd - 1);

    dtsContent = dtsContent.substring(0, itemEnd - 1) + messageCompose + dtsContent.substring(itemEnd);

    // Copy contents of MessageRead interface to Item.
    itemEnd = dtsContent.indexOf("    }", itemStart);

    let messageReadInterface = dtsContent.indexOf("interface MessageRead ");
    let messageReadStart = dtsContent.indexOf("{", messageReadInterface);
    let messageReadEnd = dtsContent.indexOf("    }", messageReadStart);
    let messageRead = dtsContent.substring(messageReadStart + 1, messageReadEnd - 1);

    dtsContent = dtsContent.substring(0, itemEnd - 1) + messageRead + dtsContent.substring(itemEnd);

    /** APPOINTMENT */
    // Initial Appointment details.
    appointmentInterface = dtsContentForAppointment.indexOf("interface Appointment ");
    appointmentStart = dtsContentForAppointment.indexOf("{", appointmentInterface);
    appointmentEnd = dtsContentForAppointment.indexOf("    }", appointmentStart);

    // Copy original contents of Item Interface to Appointment.
    itemInterface = dtsContentForAppointment.indexOf("interface Item ");
    itemStart = dtsContentForAppointment.indexOf("{", itemInterface);
    itemEnd = dtsContentForAppointment.indexOf("    }", itemStart);
    let item = dtsContentForAppointment.substring(itemStart + 1, itemEnd - 1);

    dtsContentForAppointment = dtsContentForAppointment.substring(0, appointmentEnd - 1) + item + dtsContentForAppointment.substring(appointmentEnd);

    // Copy contents of AppointmentCompose to Appointment.
    /*appointmentEnd = dtsContentForAppointment.indexOf("    }", appointmentStart);

    appointmentComposeInterface = dtsContentForAppointment.indexOf("interface AppointmentCompose ");
    appointmentComposeStart = dtsContentForAppointment.indexOf("{", appointmentComposeInterface);
    appointmentComposeEnd = dtsContentForAppointment.indexOf("    }", appointmentComposeStart);
    appointmentCompose = dtsContentForAppointment.substring(appointmentComposeStart + 1, appointmentComposeEnd - 1);

    dtsContentForAppointment = dtsContentForAppointment.substring(0, appointmentEnd - 1) + appointmentCompose + dtsContentForAppointment.substring(appointmentEnd);

    // Copy contents of AppointmentRead to Appointment.
    appointmentEnd = dtsContentForAppointment.indexOf("    }", appointmentStart);

    appointmentReadInterface = dtsContentForAppointment.indexOf("interface AppointmentRead ");
    appointmentReadStart = dtsContentForAppointment.indexOf("{", appointmentReadInterface);
    appointmentReadEnd = dtsContentForAppointment.indexOf("    }", appointmentReadStart);
    appointmentRead = dtsContentForAppointment.substring(appointmentReadStart + 1, appointmentReadEnd - 1);

    dtsContentForAppointment = dtsContentForAppointment.substring(0, appointmentEnd - 1) + appointmentRead + dtsContentForAppointment.substring(appointmentEnd);*/

    // Copy updated Appointment to dtsContent.
    let appointmentInterfaceFinal = dtsContentForAppointment.indexOf("interface Appointment ");
    let appointmentStartFinal = dtsContentForAppointment.indexOf("{", appointmentInterfaceFinal);
    let appointmentEndFinal = dtsContentForAppointment.indexOf("    }", appointmentStartFinal);
    let appointmentFinal = dtsContentForAppointment.substring(appointmentStartFinal + 1, appointmentEndFinal - 1);

    appointmentInterface = dtsContent.indexOf("interface Appointment ");
    appointmentStart = dtsContent.indexOf("{", appointmentInterface);
    appointmentEnd = dtsContent.indexOf("    }", appointmentStart);

    dtsContent = dtsContent.substring(0, appointmentStart + 1) + appointmentFinal + dtsContent.substring(appointmentEnd);

    /** APPOINTMENTCOMPOSE */
    // Initial AppointmentCompose details.
    appointmentComposeInterface = dtsContentForAppointmentCompose.indexOf("interface AppointmentCompose ");
    appointmentComposeStart = dtsContentForAppointmentCompose.indexOf("{", appointmentComposeInterface);
    appointmentComposeEnd = dtsContentForAppointmentCompose.indexOf("    }", appointmentComposeStart);

    // Copy original contents of Item Interface to AppointmentCompose.
    itemInterface = dtsContentForAppointmentCompose.indexOf("interface Item ");
    itemStart = dtsContentForAppointmentCompose.indexOf("{", itemInterface);
    itemEnd = dtsContentForAppointmentCompose.indexOf("    }", itemStart);
    item = dtsContentForAppointmentCompose.substring(itemStart + 1, itemEnd - 1);

    dtsContentForAppointmentCompose = dtsContentForAppointmentCompose.substring(0, appointmentComposeEnd - 1) + item + dtsContentForAppointmentCompose.substring(appointmentComposeEnd);

    // Copy contents of Appointment to AppointmentCompose.
    appointmentComposeEnd = dtsContentForAppointmentCompose.indexOf("    }", appointmentComposeStart);

    appointmentInterface = dtsContentForAppointmentCompose.indexOf("interface Appointment ");
    appointmentStart = dtsContentForAppointmentCompose.indexOf("{", appointmentInterface);
    appointmentEnd = dtsContentForAppointmentCompose.indexOf("    }", appointmentStart);
    appointment = dtsContentForAppointmentCompose.substring(appointmentStart + 1, appointmentEnd - 1);

    dtsContentForAppointmentCompose = dtsContentForAppointmentCompose.substring(0, appointmentComposeEnd - 1) + appointment + dtsContentForAppointmentCompose.substring(appointmentComposeEnd);

    // Copy contents of ItemCompose to AppointmentCompose.
    appointmentComposeEnd = dtsContentForAppointmentCompose.indexOf("    }", appointmentComposeStart);

    itemComposeInterface = dtsContentForAppointmentCompose.indexOf("interface ItemCompose ");
    itemComposeStart = dtsContentForAppointmentCompose.indexOf("{", itemComposeInterface);
    itemComposeEnd = dtsContentForAppointmentCompose.indexOf("    }", itemComposeStart);
    itemCompose = dtsContentForAppointmentCompose.substring(itemComposeStart + 1, itemComposeEnd - 1);

    dtsContentForAppointmentCompose = dtsContentForAppointmentCompose.substring(0, appointmentComposeEnd - 1) + itemCompose + dtsContentForAppointmentCompose.substring(appointmentComposeEnd);

    // Copy updated AppointmentCompose to dtsContent.
    let appointmentComposeInterfaceFinal = dtsContentForAppointmentCompose.indexOf("interface AppointmentCompose ");
    let appointmentComposeStartFinal = dtsContentForAppointmentCompose.indexOf("{", appointmentComposeInterfaceFinal);
    let appointmentComposeEndFinal = dtsContentForAppointmentCompose.indexOf("    }", appointmentComposeStartFinal);
    let appointmentComposeFinal = dtsContentForAppointmentCompose.substring(appointmentComposeStartFinal + 1, appointmentComposeEndFinal - 1);

    appointmentComposeInterface = dtsContent.indexOf("interface AppointmentCompose ");
    appointmentComposeStart = dtsContent.indexOf("{", appointmentComposeInterface);
    appointmentComposeEnd = dtsContent.indexOf("    }", appointmentComposeStart);

    dtsContent = dtsContent.substring(0, appointmentComposeStart + 1) + appointmentComposeFinal + dtsContent.substring(appointmentComposeEnd);

    /** APPOINTMENTREAD */
    // Initial AppointmentRead details.
    appointmentReadInterface = dtsContentForAppointmentRead.indexOf("interface AppointmentRead ");
    appointmentReadStart = dtsContentForAppointmentRead.indexOf("{", appointmentReadInterface);
    appointmentReadEnd = dtsContentForAppointmentRead.indexOf("    }", appointmentReadStart);

    // Copy original contents of Item Interface to AppointmentRead.
    itemInterface = dtsContentForAppointmentRead.indexOf("interface Item ");
    itemStart = dtsContentForAppointmentRead.indexOf("{", itemInterface);
    itemEnd = dtsContentForAppointmentRead.indexOf("    }", itemStart);
    item = dtsContentForAppointmentRead.substring(itemStart + 1, itemEnd - 1);

    dtsContentForAppointmentRead = dtsContentForAppointmentRead.substring(0, appointmentReadEnd - 1) + item + dtsContentForAppointmentRead.substring(appointmentReadEnd);

    // Copy contents of Appointment to AppointmentRead.
    appointmentReadEnd = dtsContentForAppointmentRead.indexOf("    }", appointmentReadStart);

    appointmentInterface = dtsContentForAppointmentRead.indexOf("interface Appointment ");
    appointmentStart = dtsContentForAppointmentRead.indexOf("{", appointmentInterface);
    appointmentEnd = dtsContentForAppointmentRead.indexOf("    }", appointmentStart);
    appointment = dtsContentForAppointmentRead.substring(appointmentStart + 1, appointmentEnd - 1);

    dtsContentForAppointmentRead = dtsContentForAppointmentRead.substring(0, appointmentReadEnd - 1) + appointment + dtsContentForAppointmentRead.substring(appointmentReadEnd);

    // Copy contents of ItemRead to AppointmentRead.
    appointmentReadEnd = dtsContentForAppointmentRead.indexOf("    }", appointmentReadStart);

    itemReadInterface = dtsContentForAppointmentRead.indexOf("interface ItemRead ");
    itemReadStart = dtsContentForAppointmentRead.indexOf("{", itemReadInterface);
    itemReadEnd = dtsContentForAppointmentRead.indexOf("    }", itemReadStart);
    itemRead = dtsContentForAppointmentRead.substring(itemReadStart + 1, itemReadEnd - 1);

    dtsContentForAppointmentRead = dtsContentForAppointmentRead.substring(0, appointmentReadEnd - 1) + itemRead + dtsContentForAppointmentRead.substring(appointmentReadEnd);

    // Copy updated AppointmentRead to dtsContent.
    let appointmentReadInterfaceFinal = dtsContentForAppointmentRead.indexOf("interface AppointmentRead ");
    let appointmentReadStartFinal = dtsContentForAppointmentRead.indexOf("{", appointmentReadInterfaceFinal);
    let appointmentReadEndFinal = dtsContentForAppointmentRead.indexOf("    }", appointmentReadStartFinal);
    let appointmentReadFinal = dtsContentForAppointmentRead.substring(appointmentReadStartFinal + 1, appointmentReadEndFinal - 1);

    appointmentReadInterface = dtsContent.indexOf("interface AppointmentRead ");
    appointmentReadStart = dtsContent.indexOf("{", appointmentReadInterface);
    appointmentReadEnd = dtsContent.indexOf("    }", appointmentReadStart);

    dtsContent = dtsContent.substring(0, appointmentReadStart + 1) + appointmentReadFinal + dtsContent.substring(appointmentReadEnd);

    /** MESSAGE */
    // Initial Message details.
    messageInterface = dtsContentForMessage.indexOf("interface Message ");
    messageStart = dtsContentForMessage.indexOf("{", messageInterface);
    messageEnd = dtsContentForMessage.indexOf("    }", messageStart);

    // Copy original contents of Item Interface to Message.
    itemInterface = dtsContentForMessage.indexOf("interface Item ");
    itemStart = dtsContentForMessage.indexOf("{", itemInterface);
    itemEnd = dtsContentForMessage.indexOf("    }", itemStart);
    item = dtsContentForMessage.substring(itemStart + 1, itemEnd - 1);

    dtsContentForMessage = dtsContentForMessage.substring(0, messageEnd - 1) + item + dtsContentForMessage.substring(messageEnd);

    // Copy contents of MessageCompose to Message.
    /*messageEnd = dtsContentForMessage.indexOf("    }", messageStart);

    messageComposeInterface = dtsContentForMessage.indexOf("interface MessageCompose ");
    messageComposeStart = dtsContentForMessage.indexOf("{", messageComposeInterface);
    messageComposeEnd = dtsContentForMessage.indexOf("    }", messageComposeStart);
    messageCompose = dtsContentForMessage.substring(messageComposeStart + 1, messageComposeEnd - 1);

    dtsContentForMessage = dtsContentForMessage.substring(0, messageEnd - 1) + messageCompose + dtsContentForMessage.substring(messageEnd);

    // Copy contents of MessageRead to Message.
    messageEnd = dtsContentForMessage.indexOf("    }", messageStart);

    messageReadInterface = dtsContentForMessage.indexOf("interface MessageRead ");
    messageReadStart = dtsContentForMessage.indexOf("{", messageReadInterface);
    messageReadEnd = dtsContentForMessage.indexOf("    }", messageReadStart);
    messageRead = dtsContentForMessage.substring(messageReadStart + 1, messageReadEnd - 1);

    dtsContentForMessage = dtsContentForMessage.substring(0, messageEnd - 1) + messageRead + dtsContentForMessage.substring(messageEnd);*/

    // Copy updated Message to dtsContent.
    let messageInterfaceFinal = dtsContentForMessage.indexOf("interface Message ");
    let messageStartFinal = dtsContentForMessage.indexOf("{", messageInterfaceFinal);
    let messageEndFinal = dtsContentForMessage.indexOf("    }", messageStartFinal);
    let messageFinal = dtsContentForMessage.substring(messageStartFinal + 1, messageEndFinal - 1);

    messageInterface = dtsContent.indexOf("interface Message ");
    messageStart = dtsContent.indexOf("{", messageInterface);
    messageEnd = dtsContent.indexOf("    }", messageStart);

    dtsContent = dtsContent.substring(0, messageStart + 1) + messageFinal + dtsContent.substring(messageEnd);

    /** MESSAGECOMPOSE */
    // Initial MessageCompose details.
    messageComposeInterface = dtsContentForMessageCompose.indexOf("interface MessageCompose ");
    messageComposeStart = dtsContentForMessageCompose.indexOf("{", messageComposeInterface);
    messageComposeEnd = dtsContentForMessageCompose.indexOf("    }", messageComposeStart);

    // Copy original contents of Item Interface to MessageCompose.
    itemInterface = dtsContentForMessageCompose.indexOf("interface Item ");
    itemStart = dtsContentForMessageCompose.indexOf("{", itemInterface);
    itemEnd = dtsContentForMessageCompose.indexOf("    }", itemStart);
    item = dtsContentForMessageCompose.substring(itemStart + 1, itemEnd - 1);

    dtsContentForMessageCompose = dtsContentForMessageCompose.substring(0, messageComposeEnd - 1) + item + dtsContentForMessageCompose.substring(messageComposeEnd);

    // Copy contents of Message to MessageCompose.
    messageComposeEnd = dtsContentForMessageCompose.indexOf("    }", messageComposeStart);

    messageInterface = dtsContentForMessageCompose.indexOf("interface Message ");
    messageStart = dtsContentForMessageCompose.indexOf("{", messageInterface);
    messageEnd = dtsContentForMessageCompose.indexOf("    }", messageStart);
    message = dtsContentForMessageCompose.substring(messageStart + 1, messageEnd - 1);

    dtsContentForMessageCompose = dtsContentForMessageCompose.substring(0, messageComposeEnd - 1) + message + dtsContentForMessageCompose.substring(messageComposeEnd);

    // Copy contents of ItemCompose to MessageCompose.
    messageComposeEnd = dtsContentForMessageCompose.indexOf("    }", messageComposeStart);

    itemComposeInterface = dtsContentForMessageCompose.indexOf("interface ItemCompose ");
    itemComposeStart = dtsContentForMessageCompose.indexOf("{", itemComposeInterface);
    itemComposeEnd = dtsContentForMessageCompose.indexOf("    }", itemComposeStart);
    itemCompose = dtsContentForMessageCompose.substring(itemComposeStart + 1, itemComposeEnd - 1);

    dtsContentForMessageCompose = dtsContentForMessageCompose.substring(0, messageComposeEnd - 1) + itemCompose + dtsContentForMessageCompose.substring(messageComposeEnd);

    // Copy updated MessageCompose to dtsContent.
    let messageComposeInterfaceFinal = dtsContentForMessageCompose.indexOf("interface MessageCompose ");
    let messageComposeStartFinal = dtsContentForMessageCompose.indexOf("{", messageComposeInterfaceFinal);
    let messageComposeEndFinal = dtsContentForMessageCompose.indexOf("    }", messageComposeStartFinal);
    let messageComposeFinal = dtsContentForMessageCompose.substring(messageComposeStartFinal + 1, messageComposeEndFinal - 1);

    messageComposeInterface = dtsContent.indexOf("interface MessageCompose ");
    messageComposeStart = dtsContent.indexOf("{", messageComposeInterface);
    messageComposeEnd = dtsContent.indexOf("    }", messageComposeStart);

    dtsContent = dtsContent.substring(0, messageComposeStart + 1) + messageComposeFinal + dtsContent.substring(messageComposeEnd);

    /** MESSAGEREAD */
    // Initial MessageRead details.
    messageReadInterface = dtsContentForMessageRead.indexOf("interface MessageRead ");
    messageReadStart = dtsContentForMessageRead.indexOf("{", messageReadInterface);
    messageReadEnd = dtsContentForMessageRead.indexOf("    }", messageReadStart);

    // Copy original contents of Item Interface to MessageRead.
    itemInterface = dtsContentForMessageRead.indexOf("interface Item ");
    itemStart = dtsContentForMessageRead.indexOf("{", itemInterface);
    itemEnd = dtsContentForMessageRead.indexOf("    }", itemStart);
    item = dtsContentForMessageRead.substring(itemStart + 1, itemEnd - 1);

    dtsContentForMessageRead = dtsContentForMessageRead.substring(0, messageReadEnd - 1) + item + dtsContentForMessageRead.substring(messageReadEnd);

    // Copy contents of Message to MessageRead.
    messageReadEnd = dtsContentForMessageRead.indexOf("    }", messageReadStart);

    messageInterface = dtsContentForMessageRead.indexOf("interface Message ");
    messageStart = dtsContentForMessageRead.indexOf("{", messageInterface);
    messageEnd = dtsContentForMessageRead.indexOf("    }", messageStart);
    message = dtsContentForMessageRead.substring(messageStart + 1, messageEnd - 1);

    dtsContentForMessageRead = dtsContentForMessageRead.substring(0, messageReadEnd - 1) + message + dtsContentForMessageRead.substring(messageReadEnd);

    // Copy contents of ItemRead to MessageRead.
    messageReadEnd = dtsContentForMessageRead.indexOf("    }", messageReadStart);

    itemReadInterface = dtsContentForMessageRead.indexOf("interface ItemRead ");
    itemReadStart = dtsContentForMessageRead.indexOf("{", itemReadInterface);
    itemReadEnd = dtsContentForMessageRead.indexOf("    }", itemReadStart);
    itemRead = dtsContentForMessageRead.substring(itemReadStart + 1, itemReadEnd - 1);

    dtsContentForMessageRead = dtsContentForMessageRead.substring(0, messageReadEnd - 1) + itemRead + dtsContentForMessageRead.substring(messageReadEnd);

    // Copy updated MessageRead to dtsContent.
    let messageReadInterfaceFinal = dtsContentForMessageRead.indexOf("interface MessageRead ");
    let messageReadStartFinal = dtsContentForMessageRead.indexOf("{", messageReadInterfaceFinal);
    let messageReadEndFinal = dtsContentForMessageRead.indexOf("    }", messageReadStartFinal);
    let messageReadFinal = dtsContentForMessageRead.substring(messageReadStartFinal + 1, messageReadEndFinal - 1);

    messageReadInterface = dtsContent.indexOf("interface MessageRead ");
    messageReadStart = dtsContent.indexOf("{", messageReadInterface);
    messageReadEnd = dtsContent.indexOf("    }", messageReadStart);

    dtsContent = dtsContent.substring(0, messageReadStart + 1) + messageReadFinal + dtsContent.substring(messageReadEnd);

    /** ITEMCOMPOSE */
    // Initial ItemCompose details.
    itemComposeInterface = dtsContentForItemCompose.indexOf("interface ItemCompose ");
    itemComposeStart = dtsContentForItemCompose.indexOf("{", itemComposeInterface);
    itemComposeEnd = dtsContentForItemCompose.indexOf("    }", itemComposeStart);

    // Copy original contents of Item Interface to ItemCompose.
    itemInterface = dtsContentForItemCompose.indexOf("interface Item ");
    itemStart = dtsContentForItemCompose.indexOf("{", itemInterface);
    itemEnd = dtsContentForItemCompose.indexOf("    }", itemStart);
    item = dtsContentForItemCompose.substring(itemStart + 1, itemEnd - 1);

    dtsContentForItemCompose = dtsContentForItemCompose.substring(0, itemComposeEnd - 1) + item + dtsContentForItemCompose.substring(itemComposeEnd);

    // Copy contents of AppointmentCompose to ItemCompose.
    /*itemComposeEnd = dtsContentForItemCompose.indexOf("    }", itemComposeStart);

    appointmentComposeInterface = dtsContentForItemCompose.indexOf("interface AppointmentCompose ");
    appointmentComposeStart = dtsContentForItemCompose.indexOf("{", appointmentComposeInterface);
    appointmentComposeEnd = dtsContentForItemCompose.indexOf("    }", appointmentComposeStart);
    appointmentCompose = dtsContentForItemCompose.substring(appointmentComposeStart + 1, appointmentComposeEnd - 1);

    dtsContentForItemCompose = dtsContentForItemCompose.substring(0, itemComposeEnd - 1) + appointmentCompose + dtsContentForItemCompose.substring(itemComposeEnd);

    // Copy contents of MessageCompose to ItemCompose.
    itemComposeEnd = dtsContentForItemCompose.indexOf("    }", itemComposeStart);

    messageComposeInterface = dtsContentForItemCompose.indexOf("interface MessageCompose ");
    messageComposeStart = dtsContentForItemCompose.indexOf("{", messageComposeInterface);
    messageComposeEnd = dtsContentForItemCompose.indexOf("    }", messageComposeStart);
    messageCompose = dtsContentForItemCompose.substring(messageComposeStart + 1, messageComposeEnd - 1);

    dtsContentForItemCompose = dtsContentForItemCompose.substring(0, itemComposeEnd - 1) + messageCompose + dtsContentForItemCompose.substring(itemComposeEnd);*/

    // Copy updated ItemCompose to dtsContent.
    let itemComposeInterfaceFinal = dtsContentForItemCompose.indexOf("interface ItemCompose ");
    let itemComposeStartFinal = dtsContentForItemCompose.indexOf("{", itemComposeInterfaceFinal);
    let itemComposeEndFinal = dtsContentForItemCompose.indexOf("    }", itemComposeStartFinal);
    let itemComposeFinal = dtsContentForItemCompose.substring(itemComposeStartFinal + 1, itemComposeEndFinal - 1);

    itemComposeInterface = dtsContent.indexOf("interface ItemCompose ");
    itemComposeStart = dtsContent.indexOf("{", itemComposeInterface);
    itemComposeEnd = dtsContent.indexOf("    }", itemComposeStart);

    dtsContent = dtsContent.substring(0, itemComposeStart + 1) + itemComposeFinal + dtsContent.substring(itemComposeEnd);

    /** ITEMREAD */
    // Initial ItemRead details.
    itemReadInterface = dtsContentForItemRead.indexOf("interface ItemRead ");
    itemReadStart = dtsContentForItemRead.indexOf("{", itemReadInterface);
    itemReadEnd = dtsContentForItemRead.indexOf("    }", itemReadStart);

    // Copy original contents of Item Interface to ItemRead.
    itemInterface = dtsContentForItemRead.indexOf("interface Item ");
    itemStart = dtsContentForItemRead.indexOf("{", itemInterface);
    itemEnd = dtsContentForItemRead.indexOf("    }", itemStart);
    item = dtsContentForItemRead.substring(itemStart + 1, itemEnd - 1);

    dtsContentForItemRead = dtsContentForItemRead.substring(0, itemReadEnd - 1) + item + dtsContentForItemRead.substring(itemReadEnd);

    // Copy contents of AppointmentRead to ItemRead.
    /*itemReadEnd = dtsContentForItemRead.indexOf("    }", itemReadStart);

    appointmentReadInterface = dtsContentForItemRead.indexOf("interface AppointmentRead ");
    appointmentReadStart = dtsContentForItemRead.indexOf("{", appointmentReadInterface);
    appointmentReadEnd = dtsContentForItemRead.indexOf("    }", appointmentReadStart);
    appointmentRead = dtsContentForItemRead.substring(appointmentReadStart + 1, appointmentReadEnd - 1);

    dtsContentForItemRead = dtsContentForItemRead.substring(0, itemReadEnd - 1) + appointmentRead + dtsContentForItemRead.substring(itemReadEnd);

    // Copy contents of MessageRead to ItemRead.
    itemReadEnd = dtsContentForItemRead.indexOf("    }", itemReadStart);

    messageReadInterface = dtsContentForItemRead.indexOf("interface MessageRead ");
    messageReadStart = dtsContentForItemRead.indexOf("{", messageReadInterface);
    messageReadEnd = dtsContentForItemRead.indexOf("    }", messageReadStart);
    messageRead = dtsContentForItemRead.substring(messageReadStart + 1, messageReadEnd - 1);

    dtsContentForItemRead = dtsContentForItemRead.substring(0, itemReadEnd - 1) + messageRead + dtsContentForItemRead.substring(itemReadEnd);*/

    // Copy updated ItemRead to dtsContent.
    let itemReadInterfaceFinal = dtsContentForItemRead.indexOf("interface ItemRead ");
    let itemReadStartFinal = dtsContentForItemRead.indexOf("{", itemReadInterfaceFinal);
    let itemReadEndFinal = dtsContentForItemRead.indexOf("    }", itemReadStartFinal);
    let itemReadFinal = dtsContentForItemRead.substring(itemReadStartFinal + 1, itemReadEndFinal - 1);

    itemReadInterface = dtsContent.indexOf("interface ItemRead ");
    itemReadStart = dtsContent.indexOf("{", itemReadInterface);
    itemReadEnd = dtsContent.indexOf("    }", itemReadStart);

    dtsContent = dtsContent.substring(0, itemReadStart + 1) + itemReadFinal + dtsContent.substring(itemReadEnd);


    return dtsContent;
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
        .replace(/([ ]*)load\(option\?: string \| string\[\]\): (Excel|Word|OneNote|Visio)\.(.*);/g,
                 "$1/**\n$1 * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.\n$1 * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.\n$1 */\n$1load(propertyNames?: string | string[]): $2.$3;")
        .replace(/([ ]*)load\(option\?: {\n[ ]*select\?: string;\n[ ]*expand\?: string;\n[ ]*}\): (Excel|Word|OneNote|Visio)\.(.*);/gm,
                 "$1/**\n$1 * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.\n$1 * @param propertyNamesAndPaths - Where propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.\n$1 */\n$1load(propertyNamesAndPaths?: { select?: string; expand?: string; }): $2.$3;")
        .replace(/([ ]*)load\(option\?: (Excel|Word|OneNote|Visio)\.Interfaces\.(.*)CollectionLoadOptions & [Excel|Word|OneNote|Visio]\.Interfaces\.CollectionLoadOptions\): [Excel|Word|OneNote|Visio]\.[.*]Collection;/g,
                 "$1/**\n$1 * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.\n$1 * @param collectionLoadOptions - Where collectionLoadOptions.select is a comma-delimited string that specifies the properties to load, and collectionLoadOptions.expand is a comma-delimited string that specifies the navigation properties to load. collectionLoadOptions.top specifies the maximum number of collection items that can be included in the result. collectionLoadOptions.skip specifies the number of items that are to be skipped and not included in the result. If collectionLoadOptions.top is specified, the result set will start after skipping the specified number of items.\n$1 */\n$1load(collectionLoadOptions?: $2.Interfaces.$3CollectionLoadOptions & $2.Interfaces.CollectionLoadOptions): $2.$3Collection;")
        .replace(/(extends OfficeCore.RequestContext)/g, `extends OfficeExtension.ClientRequestContext`));
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

    return dtsString.replace(/([a-zA-Z]+)(\??: \"[a-zA-Z]*\".*:)/g, "$1String$2");
}

async function tryCatch(call: () => Promise<void>) {
    try {
        await call();
    } catch (e) {
        console.error(e);
        process.exit(1);
    }
}
