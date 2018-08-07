#!/usr/bin/env node --harmony

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
    let rootPushed = false;

    // look for existing folders to move
    let outlookFolders : string[] = ["MailboxEnums"];

    // create folders for Excel subcategories
    let excelEnumRoot = {"name": "Enums", "uid": "", "items": [] as any};
    let excelEnumFilter : string [] = ["BindingType", "BorderIndex", "BorderLineStyle", "BorderWeight", "BuiltInStyle", "CalculationMode", "CalculationType", "ChartAxisCategoryType", "ChartAxisDisplayUnit", "ChartAxisGroup", "ChartAxisPosition", "ChartAxisScaleType", "ChartAxisTickLabelPosition", "ChartAxisTickMark", "ChartAxisTimeUnit", "ChartAxisType", "ChartDataLabelPosition", "ChartLegendPosition", "ChartLineStyle", "ChartMarkerStyle", "ChartSeriesBy", "ChartTextHorizontalAlignment", "ChartTextVerticalAlignment", "ChartTitlePosition", "ChartTrendlineType", "ChartType", "ChartUnderlineStyle", "ClearApplyTo", "ConditionalCellValueOperator", "ConditionalDataBarAxisFormat", "ConditionalDataBarDirection", "ConditionalFormatColorCriterionType", "ConditionalFormatDirection", "ConditionalFormatIconRuleType", "ConditionalFormatPresetCriterion", "ConditionalFormatRuleType", "ConditionalFormatType", "ConditionalIconCriterionOperator", "ConditionalRangeBorderIndex", "ConditionalRangeBorderLineStyle", "ConditionalRangeFontUnderlineStyle", "ConditionalTextOperator", "ConditionalTopBottomCriterionType", "DataChangeType", "DeleteShiftDirection", "DocumentPropertyItem", "DocumentPropertyType", "DynamicFilterCriteria", "ErrorCodes", "EventSource", "EventType", "FilterDatetimeSpecificity", "FilterOn", "FilterOperator", "HorizontalAlignment", "IconSet", "ImageFittingMode", "InsertShiftDirection", "NamedItemScope", "NamedItemType", "PageOrientation", "ProtectionSelectionMode", "RangeUnderlineStyle", "RangeValueType", "ReadingOrder", "SheetVisibility", "SortDataOption", "SortMethod", "SortOn", "SortOrientation", "VerticalAlignment", "WorksheetPositionType"];
    let excelEventArgsFilter : string [] = ["BindingDataChangedEventArgs", "BindingSelectionChangedEventArgs", "SelectionChangedEventArgs", "SettingsChangedEventArgs", "TableChangedEventArgs", "TableSelectionChangedEventArgs", "WorksheetActivatedEventArgs", "WorksheetAddedEventArgs", "WorksheetChangedEventArgs", "WorksheetDeactivatedEventArgs", "WorksheetDeletedEventArgs", "WorksheetSelectionChangedEventArgs"];
    let excelIconSetFilter : string [] = ["FiveArrowsGraySet", "FiveArrowsSet", "FiveBoxesSet", "FiveQuartersSet", "FiveRatingSet", "FourArrowsGraySet", "FourArrowsSet", "FourRatingSet", "FourRedToBlackSet", "FourTrafficLightsSet", "IconCollections", "ThreeArrowsGraySet", "ThreeArrowsSet", "ThreeFlagsSet",  "ThreeSignsSet", "ThreeStarsSet",  "ThreeSymbols2Set", "ThreeSymbolsSet", "ThreeTrafficLights1Set", "ThreeTrafficLights2Set", "ThreeTrianglesSet"];
    let excelInterfaceFilter : string [] = ["ConditionalCellValueRule", "ConditionalCellValueRule", "ConditionalColorScaleCriteria", "ConditionalColorScaleCriterion", "ConditionalDataBarRule", "ConditionalIconCriterion", "ConditionalPresetCriteriaRule", "ConditionalTextComparisonRule", "ConditionalTextComparisonRule", "ConditionalTopBottomRule", "FilterCrieteria", "FilterDatetime", "Icon", "IconCollections", "RangeHyperlink", "RangeReference", "RunOptions", "SortField", "WorksheetProtectionOptions"];

    // create folders for OneNote subcategories
    let oneNoteEnumRoot = {"name": "Enums", "uid": "", "items": [] as any};
    let oneNoteEnumFilter : string [] = ["EntityType", "ErrorCodes", "InsertLocation", "ListType", "NoteTagStatus", "NoteTagType", "NumberType", "PageContentType", "ParagraphType"];
    let oneNoteInterfaceFilter : string[] = ["ImageOcrData", "InkStrokePointer", "ParagraphInfo"];

    // create folders for word subcategories
    let wordEnumRoot = {"name": "Enums", "uid": "", "items": [] as any};
    let wordEnumFilter : string [] = ["Alignment", "BodyType", "BorderLocation", "BorderType", "BreakType", "CellPaddingLocation", "ContentControlAppearance", "ContentControlType", "DocumentPropertyType", "ErrorCodes", "FileContentFormat", "HeaderFooterType", "ImageFormat", "InsertLocation", "ListBullet", "ListLevelType", "ListNumbering", "LocationRelation", "RangeLocation", "SelectionMode", "Style", "TapObjectType", "UnderlineType", "VerticalAlignment"];

    // create folders for shared API subcategories
    let sharedEnumRoot = {"name": "Enums", "uid": "", "items": [] as any};
    let sharedEnumFilter : string [] = ["ActiveView", "AsyncResultStatus", "BindingType", "CoercionType", "CustomXMLNodeType", "DocumentMode", "EventType", "FileType", "FilterType", "GoToType", "HostType", "InitializationReason", "PlatformType", "ProjectProjectFields", "ProjectResourceFields", "ProjectTaskFields", "ProjectViewTypes", "SelectionMode", "Table", "ValueFormat"];

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

    // process all packages except 'office' (Shared API)
    origToc.items.forEach((rootItem, rootIndex) => {
        rootItem.items.forEach((packageItem, packageIndex) => {
            if (packageItem.name !== 'office') {
                const packageName = packageItem.name === 'onenote' ? 'OneNote' : packageItem.name.substr(0, 1).toUpperCase() + packageItem.name.substr(1);
                if (packageItem.items.length === 1) {
                    packageItem.items.forEach((namespaceItem, namespaceIndex) => {
                        membersToMove.items = namespaceItem.items;
                    });
                    // if outlook, put in subfolders for versioning
                    if (packageName.toLocaleLowerCase().includes('outlook')) {
                        if (!rootPushed) { // add root in alphabetical order
                            newToc.items[0].items.push(outlookRoot);
                            rootPushed = true;
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
                        membersToMove.items.filter(item => {
                            return excelIconSetFilter.indexOf(item.name) >= 0;
                        });
                        membersToMove.items.filter(item => {
                            return excelEventArgsFilter.indexOf(item.name) >= 0;
                        });
                        let enumList = membersToMove.items.filter(item => {
                            return excelEnumFilter.indexOf(item.name) >= 0;
                        });
                        let primaryList = membersToMove.items.filter(item => {
                            return excelFilter.indexOf(item.name) < 0;
                        });
                        excelEnumRoot.items = enumList;

                        primaryList.unshift(excelEnumRoot);
                        newToc.items[0].items.push({
                            "name": packageName,
                            "uid": packageItem.uid,
                            "items":  primaryList as any
                        });
                    } else if (packageName.toLocaleLowerCase().includes('word')) {
                        let enumList = membersToMove.items.filter(item => {
                            return wordEnumFilter.indexOf(item.name) >= 0;
                        });
                        let primaryList = membersToMove.items.filter(item => {
                            return wordFilter.indexOf(item.name) < 0;
                        });
                        wordEnumRoot.items = enumList;
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
                    } else {
                        newToc.items[0].items.push({
                            "name": packageName,
                            "uid": packageItem.uid,
                            "items": membersToMove.items
                        });
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

    // process 'office' (Shared API) package
    origToc.items.forEach((rootItem, rootIndex) => {
        rootItem.items.forEach((packageItem, packageIndex) => {
            if (packageItem.name === 'office') {
                packageItem.items.forEach((namespaceItem, namespaceIndex) => {
                    membersToMove.items = namespaceItem.items;
                });
                let enumList = membersToMove.items.filter(item => {
                    return sharedEnumFilter.indexOf(item.name) >= 0;
                });
                let primaryList = membersToMove.items.filter(item => {
                    return sharedEnumFilter.indexOf(item.name) < 0;
                });
                sharedEnumRoot.items = enumList;
                primaryList.unshift(sharedEnumRoot);
                newToc.items[0].items.push({
                    "name": 'Shared API',
                    "uid": packageItem.uid,
                    "items": primaryList as any
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

