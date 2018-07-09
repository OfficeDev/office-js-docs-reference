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

    // create a folder for Excel icon sets
    let excelIconSetRoot = {"name": "Icon Sets", "uid": "excel.IconSets", "items": [] as any};
    let excelIconSetFilter : string [] = ["FiveArrowsGraySet", "FiveArrowsSet", "FiveBoxesSet", "FiveQuartersSet", "FiveRatingSet", "FourArrowsGraySet", "FourArrowsSet", "FourRatingSet", "FourRedToBlackSet", "FourTrafficLightsSet", "IconCollections", "ThreeArrowsGraySet", "ThreeArrowsSet", "ThreeFlagsSet",  "ThreeSignsSet", "ThreeStarsSet",  "ThreeSymbols2Set", "ThreeSymbolsSet", "ThreeTrafficLights1Set", "ThreeTrafficLights2Set", "ThreeTrianglesSet"];

    // create a folder for Excel icon sets
    let excelEnumRoot = {"name": "Icon Sets", "uid": "excel.Enums", "items": [] as any};
    let excelEnumFilter : string [] = ["BindingType", "BorderIndex", "BorderLineStyle", "BorderWeight", "BuiltInStyle", "CalculationMode", "CalculationType", "ChartAxisCategoryType", "ChartAxisDisplayUnit", "ChartAxisGroup", "ChartAxisPosition", "ChartAxisScaleType", "ChartAxisTickLabelPosition", "ChartAxisTickMark", "ChartAxisTimeUnit", "ChartAxisType", "ChartDataLabelPosition", "ChartLegendPosition", "ChartLineStyle", "ChartMarkerStyle", "ChartSeriesBy", "ChartTextHorizontalAlignment", "ChartTextVerticalAlignment", "ChartTitlePosition", "ChartTrendlineType", "ChartType", "ChartUnderlineStyle", "ClearApplyTo", "ConditionalCellValueOperator", "ConditionalDataBarAxisFormat", "ConditionalDataBarDirection", "ConditionalFormatColorCriterionType", "ConditionalFormatDirection", "ConditionalFormatIconRuleType", "ConditionalFormatPresetCriterion", "ConditionalFormatRuleType", "ConditionalFormatType", "ConditionalIconCriterionOperator", "ConditionalRangeBorderIndex", "ConditionalRangeBorderLineStyle", "ConditionalRangeFontUnderlineStyle", "ConditionalTextOperator", "ConditionalTopBottomCriterionType", "DataChangeType", "DeleteShiftDirection", "DocumentPropertyItem", "DocumentPropertyType", "DynamicFilterCriteria", "ErrorCodes", "EventSource", "EventType", "FilterDatetimeSpecificity", "FilterOn", "FilterOperator", "HorizontalAlignment", "IconSet", "ImageFittingMode", "InsertShiftDirection", "NamedItemScope", "NamedItemType", "PageOrientation", "ProtectionSelectionMode", "RangeUnderlineStyle", "RangeValueType", "ReadingOrder", "SheetVisibility", "SortDataOption", "SortMethod", "SortOn", "SortOrientation", "VerticalAlignment", "WorksheetPositionType"];

    // create a folder for Excel eventArgs
    let excelEventArgsRoot = {"name": "Event Args", "uid": "excel.EventArgs", "items": [] as any};
    let excelEventArgsFilter : string [] = ["BindingDataChangedEventArgs", "BindingSelectionChangedEventArgs", "SelectionChangedEventArgs", "SettingsChangedEventArgs", "TableChangedEventArgs", "TableSelectionChangedEventArgs", "WorksheetActivatedEventArgs", "WorksheetAddedEventArgs", "WorksheetChangedEventArgs", "WorksheetDeactivatedEventArgs", "WorksheetDeletedEventArgs", "WorksheetSelectionChangedEventArgs"];


    // create filter lists for types we shouldn't expose
    let outlookFilter : string[] = ['Appointment', 'AppointmentForm', 'CoercionTypeOptions', 'Diagnostics', 'ItemCompose', 'ItemRead', 'Message', 'ReplyFormAttachment', 'ReplyFormData'];
    let excelFilter: string[] = ["Interfaces"];

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
                        let iconSetList = membersToMove.items.filter(item => {
                            return excelIconSetFilter.indexOf(item.name) >= 0;
                        });
                        let notIconSetList = membersToMove.items.filter(item => {
                            return excelIconSetFilter.indexOf(item.name) < 0;
                        });
                        let enumList = membersToMove.items.filter(item => {
                            return excelEnumFilter.indexOf(item.name) >= 0;
                        });
                        let notEnumSetList = notIconSetList.filter(item => {
                            return excelEnumFilter.indexOf(item.name) < 0;
                        });
                        let eventArgsList = membersToMove.items.filter(item => {
                            return excelEventArgsFilter.indexOf(item.name) >= 0;
                        });
                        let notEventArgsList = notEnumSetList.filter(item => {
                            return excelEventArgsFilter.indexOf(item.name) < 0;
                        });
                        let primaryList = notEventArgsList.filter(item => {
                            return excelFilter.indexOf(item.name) < 0;
                        });
                        newToc.items[0].items.push({
                            "name": packageName,
                            "uid": packageItem.uid,
                            "items":  primaryList as any
                        });
                        newToc.items[0].items[0].items.push(excelIconSetRoot);
                        excelIconSetRoot.items = iconSetList;
                        newToc.items[0].items[0].items.push(excelEnumRoot);
                        excelEnumRoot.items = enumList;
                        newToc.items[0].items[0].items.push(excelEventArgsRoot);
                        excelEventArgsRoot.items = eventArgsList;
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
                newToc.items[0].items.push({
                    "name": 'Shared API',
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

