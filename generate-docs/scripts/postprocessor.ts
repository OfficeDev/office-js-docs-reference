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

    const tocPath = path.resolve("../yaml/toc.yml");

    console.log("\nStarting postprocessor script...");

    console.log(`\nUpdating the structure of the TOC file: ${tocPath}`);

    let origToc = (jsyaml.safeLoad(fsx.readFileSync(tocPath).toString()) as IOrigToc);
    let newToc = <INewToc>{};
    let membersToMove = <IMembers>{};

    newToc.items = [{
        "name": origToc.items[0].name,
        "href": origToc.items[0].href
    }];
    newToc.items[0].items = [] as any;

    // create a root for all the Outlook versions
    let outlookRoot = {"name": "Outlook", "uid": "", "items": [] as any};
    let outlookRootPushed = false;

    // look for existing folders to move
    let outlookFolders : string[] = ["MailboxEnums"];

    // create folders for Excel subcategories
    let excelRoot = {"name": "Excel", "uid": "", "items": [] as any};
    let excelRootPushed = false;
    let excelEnumFilter = generateEnumList(fsx.readFileSync("../api-extractor-inputs-excel/excel.d.ts").toString());

    // let excelEventArgsFilter : string [] = ["BindingDataChangedEventArgs", "BindingSelectionChangedEventArgs", "ChartActivatedEventArgs", "ChartAddedEventArgs", "ChartDeactivatedEventArgs", "ChartDeletedEventArgs", "SelectionChangedEventArgs", "SettingsChangedEventArgs", "TableChangedEventArgs", "TableSelectionChangedEventArgs", "WorksheetActivatedEventArgs", "WorksheetAddedEventArgs", "WorksheetCalculatedEventArgs", "WorksheetChangedEventArgs", "WorksheetDeactivatedEventArgs", "WorksheetDeletedEventArgs", "WorksheetSelectionChangedEventArgs"];
    // let excelIconSetFilter : string [] = ["FiveArrowsGraySet", "FiveArrowsSet", "FiveBoxesSet", "FiveQuartersSet", "FiveRatingSet", "FourArrowsGraySet", "FourArrowsSet", "FourRatingSet", "FourRedToBlackSet", "FourTrafficLightsSet", "IconCollections", "ThreeArrowsGraySet", "ThreeArrowsSet", "ThreeFlagsSet",  "ThreeSignsSet", "ThreeStarsSet",  "ThreeSymbols2Set", "ThreeSymbolsSet", "ThreeTrafficLights1Set", "ThreeTrafficLights2Set", "ThreeTrianglesSet"];
    // let excelInterfaceFilter : string [] = ["CellPropertiesBorderLoadOptions", "CellPropertiesFillLoadOptions", "CellPropertiesFontLoadOptions", "CellPropertiesFormatLoadOptions", "CellPropertiesLoadOptions ", "ColumnPropertiesLoadOptions", "ConditionalCellValueRule", "ConditionalCellValueRule", "ConditionalColorScaleCriteria", "ConditionalColorScaleCriterion", "ConditionalDataBarRule", "ConditionalIconCriterion", "ConditionalPresetCriteriaRule", "ConditionalTextComparisonRule", "ConditionalTextComparisonRule", "ConditionalTopBottomRule", "FilterCrieteria", "FilterDatetime", "Icon", "IconCollections", "RangeHyperlink", "RangeReference", "RowPropertiesLoadOptions", "RunOptions", "SortField", "WorksheetProtectionOptions"];

    let customFunctionsRoot = {"name": "Custom Functions - Preview", "uid": "", "items": [] as any};

    // create folders for OneNote subcategories
    let oneNoteEnumRoot = {"name": "Enums", "uid": "", "items": [] as any};
    let oneNoteEnumFilter = generateEnumList(fsx.readFileSync("../api-extractor-inputs-onenote/onenote.d.ts").toString());
    //let oneNoteInterfaceFilter : string[] = ["ImageOcrData", "InkStrokePointer", "ParagraphInfo"];

    // create folders for word subcategories
    let wordEnumFilter = generateEnumList(fsx.readFileSync("../api-extractor-inputs-word/word.d.ts").toString());

    // create folders for common (shared) API subcategories
    let sharedEnumRoot = {"name": "Enums", "uid": "", "items": [] as any};
    let sharedEnumFilter : string [] = ["ActiveView", "AsyncResultStatus", "BindingType", "CoercionType", "CustomXMLNodeType", "DocumentMode", "EventType", "FileType", "FilterType", "GoToType", "HostType", "InitializationReason", "PlatformType", "ProjectProjectFields", "ProjectResourceFields", "ProjectTaskFields", "ProjectViewTypes", "SelectionMode", "Table", "ValueFormat"];

    // create filter lists for types we shouldn't expose
    //let outlookFilter : string[] = ['Appointment', 'AppointmentForm', 'CoercionTypeOptions', 'Diagnostics', 'ItemCompose', 'ItemRead', 'Message', 'ReplyFormAttachment', 'ReplyFormData'];
    let outlookFilter : string[] = ['Appointment', 'AppointmentForm', 'ItemCompose', 'ItemRead', 'Message'];
    outlookFilter = outlookFilter.concat(outlookFolders);
    let excelFilter: string[] = ["Interfaces"];
    //excelFilter = excelFilter.concat(excelIconSetFilter).concat(excelEnumFilter).concat(excelEventArgsFilter).concat(excelInterfaceFilter);
    excelFilter = excelFilter.concat(excelEnumFilter);
    let wordFilter: string[] = ["Interfaces"];
    wordFilter = wordFilter.concat(wordEnumFilter);
    let oneNoteFilter: string[] = ["Interfaces"];
    //oneNoteFilter = oneNoteFilter.concat(oneNoteEnumFilter).concat(oneNoteInterfaceFilter);
    oneNoteFilter = oneNoteFilter.concat(oneNoteEnumFilter);
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

                    // if outlook, put in subfolders for versioning
                    if (packageName.toLocaleLowerCase().includes('outlook')) {
                        if (!outlookRootPushed) { // add root in alphabetical order
                            newToc.items[0].items.push(outlookRoot);
                            outlookRootPushed = true;
                        }
                        let filterToCContent = membersToMove.items.filter(item => {
                            return outlookFilter.indexOf(item.name) < 0;
                        });
                        let folderIndex: number = 0;
                            while (folderIndex >= 0) {
                                folderIndex = membersToMove.items.findIndex(item => {
                                    return outlookFolders.indexOf(item.name) >= 0;
                                });
                                if (folderIndex >= 0) {
                                    filterToCContent.unshift(membersToMove.items.splice(folderIndex, 1)[0]);
                                }
                            }
                        if (packageName === 'Outlook') { // The version without a suffix is the preview version
                            outlookRoot.items.push({
                                "name": packageName + " - Preview",
                                "uid": packageItem.uid,
                                "items": filterToCContent
                            });
                        }
                        else {
                            let packageNameVersionFormated = packageName.replace('_1_', ' 1.');
                            outlookRoot.items.push({
                                "name": packageNameVersionFormated,
                                "uid": packageItem.uid,
                                "items": filterToCContent
                            });
                        }
                    } else if (packageName.toLocaleLowerCase().includes('excel')) {
                        if (!excelRootPushed) { // add root in alphabetical order
                            newToc.items[0].items.push(excelRoot);
                            excelRootPushed = true;
                        }
                        let enumList = membersToMove.items.filter(item => {
                             return excelEnumFilter.indexOf(item.name) >= 0;
                         });
                        let primaryList = membersToMove.items.filter(item => {
                            return excelFilter.indexOf(item.name) < 0;
                        });

                        let excelEnumRoot = {"name": "Enums", "uid": "", "items": enumList};
                        primaryList.unshift(excelEnumRoot);

                        if (packageName === 'Excel') { // The version without a suffix is the preview version
                            excelRoot.items.push({
                                "name": packageName + " - Preview",
                                "uid": packageItem.uid,
                                "items": primaryList
                            });
                        }
                        else {
                            let packageNameVersionFormated = packageName.replace('_r', ' - R');
                            excelRoot.items.push({
                                "name": packageNameVersionFormated,
                                "uid": packageItem.uid,
                                "items": primaryList
                            });
                        }
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
                    } else if (packageName.toLocaleLowerCase().includes('custom functions runtime')) {
                        customFunctionsRoot.items.push({
                            "name": packageName,
                            "uid": packageItem.uid,
                            "items":  membersToMove.items as any
                        });
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
    // Get the logical order: Preview, 1.6, 1.5, etc.
    outlookRoot.items.reverse();
    outlookRoot.items.unshift(outlookRoot.items.pop());
    // add custom functions packages under excel
    excelRoot.items.unshift(customFunctionsRoot);

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

    // write file
    fsx.writeFileSync(tocPath, jsyaml.safeDump(newToc));

    const docsSource = path.resolve("../yaml");
    const docsDestination = path.resolve("../../docs/docs-ref-autogen");

    console.log(`\nCopying docs output files to: ${docsDestination}`);

    // delete everything except the 'overview' folder from the /docs folder
    fsx.readdirSync(docsDestination)
        .filter(filename => filename !== "overview")
        .forEach(filename => fsx.removeSync(docsDestination + '/' + filename));

    fsx.readdirSync(docsSource)
        .filter(filename => filename.indexOf("outlook") >= 0 && filename.indexOf(".yml") < 0)
        .forEach(filename => {
            let subfolder = docsSource + '/' + filename;
            fsx.readdirSync(subfolder)
                .forEach(subfilename => {
                    fsx.writeFileSync(subfolder + '/' + subfilename, fsx.readFileSync(subfolder + '/' + subfilename).toString().replace(/CommonAPI/g, "Office"));
                });
        });
        fsx.readdirSync(docsSource)
        .filter(filename => filename.indexOf("office") >= 0 && filename.indexOf(".yml") < 0)
        .forEach(filename => {
            let subfolder = docsSource + '/' + filename;
            fsx.readdirSync(subfolder).filter(filename => filename.indexOf("context") >= 0)
                .forEach(subfilename => {
                    fsx.writeFileSync(subfolder + '/' + subfilename, fsx.readFileSync(subfolder + '/' + subfilename).toString().replace(/Outlook\.Mailbox/g, "Office.Mailbox").replace(/Outlook\.RoamingSettings/g, "Office.RoamingSettings"));
                });
        });

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

