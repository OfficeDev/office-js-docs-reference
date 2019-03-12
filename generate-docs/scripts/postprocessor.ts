#!/usr/bin/env node --harmony

import { generateEnumList } from './util';
import * as fsx from 'fs-extra';
import * as jsyaml from "js-yaml";
import * as path from "path";

interface IOrigToc {
    items: [
        {
            name: string,
            href: string,
            items: [
                {
                    name: string,
                    uid: string,
                    items: [
                        {
                            name: string,
                            items: [
                                {
                                    name: string,
                                    uid?: string
                                }
                            ]
                        }
                    ]
                }
            ]
        }
    ]
}

interface IMembers {
    items: [
        {
            name: string,
            uid?: string
        }
    ]
}

interface INewToc {
    items: [
        {
            name: string,
            href: string,
            items?: [
                {
                    name: string,
                    uid: string,
                    items?: [
                        {
                            name: string,
                            uid?: string,
                            items?: [
                                {
                                    name?: string,
                                    items?: [
                                        {
                                            name?: string,
                                            uid?: string
                                        }
                                    ]
                                }
                            ]
                        }
                    ]
                }
            ]
        }
    ]
}

tryCatch(async () => {
    console.log("\nStarting postprocessor script...");

    const docsSource = path.resolve("../yaml");
    const docsDestination = path.resolve("../../docs/docs-ref-autogen");

    console.log(`Deleting old docs at: ${docsDestination}`);
    // delete everything except the 'overview' folder from the /docs folder
    fsx.readdirSync(docsDestination)
        .filter(filename => filename !== "overview" && filename !== "images")
        .forEach(filename => fsx.removeSync(docsDestination + '/' + filename));

    // fix all the individual TOC files
    const commonTocFolder = path.resolve("../yaml/office");
    const commonToc = scrubAndWriteToc(commonTocFolder);
    const hostVersionMap = [{host: "excel", versions: 9},
                            {host: "onenote", versions: 1},
                            {host: "outlook", versions: 8},
                            {host: "powerpoint", versions: 1},
                            {host: "visio", versions: 1},
                            {host: "word", versions: 4}];

    hostVersionMap.forEach(category => {
        scrubAndWriteToc(path.resolve(`../yaml/${category.host}`), commonToc, category.host);
        for (let i = 1; i < category.versions; i++) {
            scrubAndWriteToc(path.resolve(`../yaml/${category.host}_1_${i}`), commonToc, category.host);
        }
    });

    console.log(`Namespace pass on Outlook docs`);
    // replace Outlook/CommonAPI namespace references with Office
    fsx.readdirSync(docsSource)
        .filter(filename => filename.indexOf("outlook") >= 0 && filename.indexOf(".yml") < 0)
        .forEach(filename => {
            let subfolder = docsSource + '/' + filename + "/outlook";
            fsx.readdirSync(subfolder)
                .forEach(subfilename => {
                    fsx.writeFileSync(subfolder + '/' + subfilename, fsx.readFileSync(subfolder + '/' + subfilename).toString().replace(/CommonAPI/g, "Office"));
                });
        });
    console.log(`Namespace pass on Office docs`);
    const officeFolder = docsSource + "/office/office";
    fsx.readdirSync(officeFolder)
        .forEach(filename => {
            fsx.writeFileSync(officeFolder + '/' + filename, fsx.readFileSync(officeFolder + '/' + filename).toString().replace(/Outlook\.Mailbox/g, "Office.Mailbox").replace(/Outlook\.RoamingSettings/g, "Office.RoamingSettings"));
        });

    console.log(`Fixing top href`);
    fsx.readdirSync(docsSource)
        .forEach(filename => {
            let subfolder = docsSource + '/' + filename;
            fsx.readdirSync(subfolder)
                .filter(subfilename => subfilename.indexOf("toc") >= 0)
                .forEach(subfilename => {
                    fsx.writeFileSync(subfolder + '/' + subfilename, fsx.readFileSync(subfolder + '/' + subfilename).toString().replace("~/docs-ref-autogen/overview/office.md", "api-ref-office-js.md"));
                });
        });

    // moving common TOC to its own folder
    fsx.copySync(commonTocFolder + "/toc.yml",  "../yaml/common/toc.yml");
    fsx.copySync(commonTocFolder + "/api-ref-office-js.md", "../yaml/common/api-ref-office-js.md");

    // remove to prevent build errors
    fsx.removeSync(commonTocFolder + "/toc.yml");
    fsx.removeSync(commonTocFolder + "/api-ref-office-js.md");

    // create global TOC
    let globalToc = <INewToc>{};
    globalToc.items = [{"name": "Excel", "href": "/javascript/api/excel"},
                       {"name": "OneNote", "href": "/javascript/api/onenote"},
                       {"name": "Outlook", "href": "/javascript/api/outlook"},
                       {"name": "PowerPoint", "href": "/javascript/api/powerpoint"},
                       {"name": "Visio", "href": "/javascript/api/visio"},
                       {"name": "Word", "href": "/javascript/api/word"},
                       {"name": "CommonAPI", "href": "/javascript/api/office"}] as any;
    fsx.writeFileSync(docsDestination + "/toc.yml", jsyaml.safeDump(globalToc));

    console.log(`Copying docs output files to: ${docsDestination}`);
    // copy docs output to /docs/docs-ref-autogen folder
    fsx.readdirSync(docsSource)
        .forEach(filename => {
        fsx.copySync(
            docsSource + '/' + filename,
            docsDestination + '/' + filename
        );
    });

    console.log("\nPostprocessor script complete!\n");

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

function fixToc(tocPath: string, commonToc: INewToc): INewToc {
    console.log(`Updating the structure of the TOC file: ${tocPath}`);

    let origToc = (jsyaml.safeLoad(fsx.readFileSync(tocPath).toString()) as IOrigToc);
    let newToc = <INewToc>{};
    let membersToMove = <IMembers>{};

    newToc.items = [{
        "name": origToc.items[0].name,
        "href": origToc.items[0].href
    }];
    newToc.items[0].items = [] as any;

    // look for existing folders to move
    let outlookFolders : string[] = ["MailboxEnums"];

    // create folders for Excel subcategories
    let excelEnumFilter = generateEnumList(fsx.readFileSync("../api-extractor-inputs-excel/excel.d.ts").toString());

    let excelEventArgsFilter : string [] = ["BindingDataChangedEventArgs", "BindingSelectionChangedEventArgs", "ChartActivatedEventArgs", "ChartAddedEventArgs", "ChartDeactivatedEventArgs", "ChartDeletedEventArgs", "SelectionChangedEventArgs", "SettingsChangedEventArgs", "TableChangedEventArgs", "TableSelectionChangedEventArgs", "WorksheetActivatedEventArgs", "WorksheetAddedEventArgs", "WorksheetCalculatedEventArgs", "WorksheetChangedEventArgs", "WorksheetDeactivatedEventArgs", "WorksheetDeletedEventArgs", "WorksheetSelectionChangedEventArgs"];
    let excelIconSetFilter : string [] = ["FiveArrowsGraySet", "FiveArrowsSet", "FiveBoxesSet", "FiveQuartersSet", "FiveRatingSet", "FourArrowsGraySet", "FourArrowsSet", "FourRatingSet", "FourRedToBlackSet", "FourTrafficLightsSet", "IconCollections", "ThreeArrowsGraySet", "ThreeArrowsSet", "ThreeFlagsSet",  "ThreeSignsSet", "ThreeStarsSet",  "ThreeSymbols2Set", "ThreeSymbolsSet", "ThreeTrafficLights1Set", "ThreeTrafficLights2Set", "ThreeTrianglesSet"];
    let excelInterfaceFilter : string [] = ["CellPropertiesBorderLoadOptions", "CellPropertiesFillLoadOptions", "CellPropertiesFontLoadOptions", "CellPropertiesFormatLoadOptions", "CellPropertiesLoadOptions ", "ColumnPropertiesLoadOptions", "ConditionalCellValueRule", "ConditionalCellValueRule", "ConditionalColorScaleCriteria", "ConditionalColorScaleCriterion", "ConditionalDataBarRule", "ConditionalIconCriterion", "ConditionalPresetCriteriaRule", "ConditionalTextComparisonRule", "ConditionalTextComparisonRule", "ConditionalTopBottomRule", "FilterCrieteria", "FilterDatetime", "Icon", "IconCollections", "RangeHyperlink", "RangeReference", "RowPropertiesLoadOptions", "RunOptions", "SortField", "WorksheetProtectionOptions"];

    let customFunctionsRoot = {"name": "Custom Functions - Preview", "uid": "", "items": [] as any};
    let customFunctionsRootPushed = false;

    // create folders for OneNote subcategories
    let oneNoteEnumRoot = {"name": "Enums", "uid": "", "items": [] as any};
    let oneNoteEnumFilter = generateEnumList(fsx.readFileSync("../api-extractor-inputs-onenote/onenote.d.ts").toString());
    let oneNoteInterfaceFilter : string[] = ["ImageOcrData", "InkStrokePointer", "ParagraphInfo"];

    // create folders for word subcategories
    let wordEnumFilter = generateEnumList(fsx.readFileSync("../api-extractor-inputs-word/word.d.ts").toString());

    // create filter lists for types we shouldn't expose
    let outlookFilter : string[] = ['Appointment', 'AppointmentForm', 'CoercionTypeOptions', 'Diagnostics', 'ItemCompose', 'ItemRead', 'Message', 'ReplyFormAttachment', 'ReplyFormData'];
    outlookFilter = outlookFilter.concat(outlookFolders);
    let excelFilter: string[] = ["Interfaces"];
    excelFilter = excelFilter.concat(excelIconSetFilter).concat(excelEnumFilter).concat(excelEventArgsFilter).concat(excelInterfaceFilter);
    let wordFilter: string[] = ["Interfaces"];
    wordFilter = wordFilter.concat(wordEnumFilter);
    let oneNoteFilter: string[] = ["Interfaces"];
    oneNoteFilter = oneNoteFilter.concat(oneNoteEnumFilter).concat(oneNoteInterfaceFilter);
    let visioFilter: string[] = ["Interfaces"];

    // process all packages except 'office' (Common "Shared" API)
    origToc.items.forEach((rootItem, rootIndex) => {
        rootItem.items.forEach((packageItem, packageIndex) => {
            if (packageItem.name !== 'office') {
                // fix host capitalization
                let packageName;
                if (packageItem.name === 'onenote') {
                    packageName = 'OneNote';
                } else if (packageItem.name === 'powerpoint') {
                    packageName = 'PowerPoint';
                } else {
                    packageName = (packageItem.name.substr(0, 1).toUpperCase() + packageItem.name.substr(1)).replace(/\-/g, ' ');
                }

                if (packageItem.items.length === 1) {
                    packageItem.items.forEach((namespaceItem, namespaceIndex) => {
                        membersToMove.items = namespaceItem.items;
                    });

                    if (packageName.toLocaleLowerCase().includes('outlook')) {
                        let filterToCContent = membersToMove.items.filter(item => {
                            return outlookFilter.indexOf(item.name) < 0;
                        });
                        // move MailboxEnums to top
                        let folderIndex: number = 0;
                        while (folderIndex >= 0) {
                            folderIndex = membersToMove.items.findIndex(item => {
                                return outlookFolders.indexOf(item.name) >= 0;
                            });
                            if (folderIndex >= 0) {
                                filterToCContent.unshift(membersToMove.items.splice(folderIndex, 1)[0]);
                            }
                        }

                        newToc.items[0].items.push({
                            "name": packageName,
                            "uid": packageItem.uid,
                            "items": filterToCContent as any
                        });
                    } else if (packageName.toLocaleLowerCase().includes('excel')) {
                        let enumList = membersToMove.items.filter(item => {
                             return excelEnumFilter.indexOf(item.name) >= 0;
                         });
                        let primaryList = membersToMove.items.filter(item => {
                            return excelFilter.indexOf(item.name) < 0;
                        });

                        let excelEnumRoot = {"name": "Enums", "uid": "", "items": enumList};
                        primaryList.unshift(excelEnumRoot);
                        newToc.items[0].items.push({
                            "name": packageName,
                            "uid": packageItem.uid,
                            "items": primaryList as any
                        });
                    } else if (packageName.toLocaleLowerCase().includes('word')) {
                        let enumList = membersToMove.items.filter(item => {
                            return wordEnumFilter.indexOf(item.name) >= 0;
                        });
                        let primaryList = membersToMove.items.filter(item => {
                            return wordFilter.indexOf(item.name) < 0;
                        });

                        let wordEnumRoot = {"name": "Enums", "uid": "", "items": enumList};
                        primaryList.unshift(wordEnumRoot);
                        newToc.items[0].items.push({
                            "name": packageName,
                            "uid": packageItem.uid,
                            "items":  primaryList as any
                        });
                    } else if (packageName.toLocaleLowerCase().includes('visio')) {
                        let primaryList = membersToMove.items.filter(item => {
                            return visioFilter.indexOf(item.name) < 0;
                        });
                        newToc.items[0].items.push({
                            "name": packageName,
                            "uid": packageItem.uid,
                            "items":  primaryList as any
                        });
                    } else if (packageName.toLocaleLowerCase().includes('onenote')) {
                        let enumList = membersToMove.items.filter(item => {
                            return oneNoteEnumFilter.indexOf(item.name) >= 0;
                        });
                        let primaryList = membersToMove.items.filter(item => {
                            return oneNoteFilter.indexOf(item.name) < 0;
                        });
                        oneNoteEnumRoot.items = enumList;
                        primaryList.unshift(oneNoteEnumRoot);
                        newToc.items[0].items.push({
                            "name": packageName,
                            "uid": packageItem.uid,
                            "items":  primaryList as any
                        });
                    } else if (packageName.toLocaleLowerCase().includes('office runtime')) {
                        customFunctionsRoot.items.push({
                            "name": packageName,
                            "uid": packageItem.uid,
                            "items":  membersToMove.items as any
                        });

                        if (!customFunctionsRootPushed) {
                            newToc.items[0].items.push(customFunctionsRoot);
                            customFunctionsRootPushed = true;
                        }
                    } else if (packageName.toLocaleLowerCase().includes('custom functions runtime')) {
                        customFunctionsRoot.items.push({
                            "name": packageName,
                            "uid": packageItem.uid,
                            "items":  membersToMove.items as any
                        });

                        if (!customFunctionsRootPushed) {
                            newToc.items[0].items.push(customFunctionsRoot);
                            customFunctionsRootPushed = true;
                        }
                    } else {
                        if (membersToMove.items) {
                            newToc.items[0].items.push({
                                "name": packageName,
                                "uid": packageItem.uid,
                                "items": membersToMove.items
                            });
                        } else {
                            newToc.items[0].items.push({
                                "name": packageName,
                                "uid": packageItem.uid,
                                "items": [] as any
                            });
                        }
                    }
                } else {
                    newToc.items[0].items.push({
                        "name": packageName,
                        "uid": packageItem.uid,
                        "items": packageItem.items
                    });
                }
            }
        });
    });

    // append the common API toc
    newToc.items[0].items.push(commonToc.items[0].items[0]);
    return newToc;
}

function fixCommonToc(tocPath: string): INewToc {
    console.log(`\nUpdating the structure of the TOC file: ${tocPath}`);

    let origToc = (jsyaml.safeLoad(fsx.readFileSync(tocPath).toString()) as IOrigToc);
    let newToc = <INewToc>{};

    newToc.items = [{
        "name": origToc.items[0].name,
        "href": origToc.items[0].href
    }];
    newToc.items[0].items = [] as any;

    // create folders for common (shared) API subcategories
    let sharedEnumRoot = {"name": "Enums", "uid": "", "items": [] as any};
    let sharedEnumFilter = generateEnumList(fsx.readFileSync("../api-extractor-inputs-office/office.d.ts").toString());

    // process 'office' (Common "Shared" API) package
    origToc.items.forEach((rootItem, rootIndex) => {
        rootItem.items.forEach((packageItem, packageIndex) => {
            if (packageItem.name === 'office') {
                packageItem.items.forEach((namespaceItem, namespaceIndex) => {
                    if (namespaceItem.name.toLocaleLowerCase() === 'office') {
                        let enumList = namespaceItem.items.filter(item => {
                            return sharedEnumFilter.indexOf(item.name) >= 0;
                        });
                        let primaryList = namespaceItem.items.filter(item => {
                            return sharedEnumFilter.indexOf(item.name) < 0;
                        });
                        sharedEnumRoot.items = enumList;
                        primaryList.unshift(sharedEnumRoot);
                        namespaceItem.items = primaryList as any;
                    }
                });
                newToc.items[0].items.push({
                    "name": 'Common API',
                    "uid": packageItem.uid,
                    "items": packageItem.items
                });
            }
        });
    });

    return newToc;
}

function addCrossHostTocStubs(toc: INewToc, hostName: string): void {
    const stubItems = [{"name": "Excel", "href": "/javascript/api/excel?view=excel-js-preview"},
                       {"name": "OneNote", "href": "/javascript/api/onenote?view=onenote-js-1.1"},
                       {"name": "Outlook", "href": "/javascript/api/outlook?view=outlook-js-preview"},
                       {"name": "PowerPoint", "href": "/javascript/api/powerpoint?view=powerpoint-js-1.1"},
                       {"name": "Visio", "href": "/javascript/api/visio?view=visio-js-1.1"},
                       {"name": "Word", "href": "/javascript/api/word?view=word-js-preview"}];

    stubItems
        .filter(stub => !hostName || !hostName.toLowerCase().includes(stub.name.toLowerCase()))
        .forEach((stubItem, stubIndex) => {
            toc.items[0].items.push(stubItem as any);
        });
}

function scrubAndWriteToc(versionFolder: string, commonToc?: INewToc, hostName?: string): INewToc {
    const tocPath = versionFolder + "/toc.yml";
    let latestToc;
    if (!commonToc) {
        latestToc = fixCommonToc(tocPath);
    } else {
        latestToc = fixToc(tocPath, commonToc);
    }

    addCrossHostTocStubs(latestToc, hostName);
    fsx.writeFileSync(tocPath, jsyaml.safeDump(latestToc));
    fsx.copySync("../../docs/docs-ref-autogen/overview/api-ref-office-js.md", versionFolder + "/api-ref-office-js.md");
    return latestToc;
}
