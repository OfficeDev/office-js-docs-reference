#!/usr/bin/env node --harmony

import { generateEnumList } from './util';
import * as fsx from 'fs-extra';
import * as jsyaml from "js-yaml";
import * as path from "path";

const OLDEST_EXCEL_RELEASE_WITH_CUSTOM_FUNCTIONS = 9;

interface Toc {
    items: [
        {
            name: string,
            href?: string,
            items: [
                {
                    name: string,
                    uid: string,
                    items: [
                        {
                            name: string,
                            uid?: string,
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

tryCatch(async () => {
    console.log("\nStarting postprocessor script...");

    const docsSource = path.resolve("../yaml");
    const docsDestination = path.resolve("../../docs/docs-ref-autogen");

    console.log(`Deleting old docs at: ${docsDestination}`);
    // delete everything except the 'overview' folder from the /docs folder
    fsx.readdirSync(docsDestination)
        .filter(filename => filename !== "overview" && filename !== "images")
        .forEach(filename => fsx.removeSync(docsDestination + '/' + filename));

    console.log(`Creating global TOC`);
    let globalToc = <Toc>{items: [{"name": "API reference"}]};
    globalToc.items[0].items = [{"name": "API reference overview", "href": "/javascript/api/overview"},
                                {"name": "Excel", "href": "/javascript/api/excel?view=excel-js-preview"},
                                {"name": "OneNote", "href": "/javascript/api/onenote?view=onenote-js-1.1"},
                                {"name": "Outlook", "href": "/javascript/api/outlook?view=outlook-js-preview"},
                                {"name": "PowerPoint", "href": "/javascript/api/powerpoint?view=powerpoint-js-1.1"},
                                {"name": "Visio", "href": "/javascript/api/visio?view=visio-js-1.1"},
                                {"name": "Word", "href": "/javascript/api/word?view=word-js-preview"},
                                {"name": "Common APIs", "href": "/javascript/api/office?view=common-js"}] as any;
    fsx.writeFileSync(docsDestination + "/toc.yml", jsyaml.safeDump(globalToc));
    fsx.writeFileSync(docsDestination + "/overview/toc.yml", jsyaml.safeDump(globalToc));

    console.log(`Copying docs output files to: ${docsDestination}`);
    // copy docs output to /docs/docs-ref-autogen folder
    fsx.readdirSync(docsSource)
        .forEach(filename => {
        fsx.copySync(
            docsSource + '/' + filename,
            docsDestination + '/' + filename
        );
    });

    // fix all the individual TOC files
    const commonToc = scrubAndWriteToc(docsDestination + "/office");
    const hostVersionMap = [{host: "excel", versions: 13}, /*not including online*/
                            {host: "onenote", versions: 1},
                            {host: "outlook", versions: 10},
                            {host: "powerpoint", versions: 2},
                            {host: "visio", versions: 1},
                            {host: "word", versions: 4}];

    hostVersionMap.forEach(category => {
        scrubAndWriteToc(path.resolve(`${docsDestination}/${category.host}`), commonToc, category.host, category.versions);
        for (let i = 1; i < category.versions; i++) {
            scrubAndWriteToc(path.resolve(`${docsDestination}/${category.host}_1_${i}`), commonToc, category.host, i);
        }
    });

    // Special case for ExcelApi Online
    scrubAndWriteToc(path.resolve(`${docsDestination}/excel_online`), commonToc, "excel", 99);


    console.log(`Namespace pass on Outlook docs`);
    // replace Outlook/CommonAPI namespace references with Office
    fsx.readdirSync(docsDestination)
        .filter(filename => filename.indexOf("outlook") >= 0 && filename.indexOf(".yml") < 0)
        .forEach(filename => {
            let subfolder = docsDestination + '/' + filename + "/outlook";
            fsx.readdirSync(subfolder)
                .forEach(subfilename => {
                    fsx.writeFileSync(subfolder + '/' + subfilename, fsx.readFileSync(subfolder + '/' + subfilename).toString().replace(/CommonAPI/g, "Office"));
                });
        });
    console.log(`Namespace pass on Office docs`);
    const officeFolder = docsDestination + "/office/office";
    fsx.readdirSync(officeFolder)
        .forEach(filename => {
            fsx.writeFileSync(officeFolder + '/' + filename, fsx.readFileSync(officeFolder + '/' + filename).toString().replace(/Outlook\.Mailbox/g, "Office.Mailbox").replace(/Outlook\.RoamingSettings/g, "Office.RoamingSettings"));
        });

    console.log(`Custom Functions API requirement set link pass`);
    fsx.readdirSync(docsDestination)
        .filter(filename => filename.indexOf("excel") >= 0 && filename.indexOf(".yml") < 0)
        .forEach(filename => {
            let subfolder = docsDestination + '/' + filename + "/custom-functions-runtime";
            if (fsx.existsSync(subfolder)) {
                fsx.readdirSync(subfolder)
                    .forEach(subfilename => {
                        fsx.writeFileSync(subfolder + '/' + subfilename,
                            fsx.readFileSync(subfolder + '/' + subfilename).toString()
                                .replace(/\/office\/dev\/add-ins\/reference\/javascript-api-for-office/g, "/office/dev/add-ins/excel/custom-functions-requirement-sets")
                                .replace(/\/office\/dev\/add-ins\/reference\/overview\/visio-javascript-reference-overview/g, "/office/dev/add-ins/excel/custom-functions-requirement-sets"));
                    });
            }
        });

    console.log(`PowerPoint API requirement set link pass`);
    fsx.readdirSync(docsDestination)
        .filter(filename => filename.indexOf("powerpoint") >= 0 && filename.indexOf(".yml") < 0)
        .forEach(filename => {
            let subfolder = docsDestination + '/' + filename + "/powerpoint";
            if (fsx.existsSync(subfolder)) {
                fsx.readdirSync(subfolder)
                    .forEach(subfilename => {
                        fsx.writeFileSync(subfolder + '/' + subfilename,
                            fsx.readFileSync(subfolder + '/' + subfilename).toString()
                                .replace(/\/office\/dev\/add-ins\/reference\/javascript-api-for-office/g, "/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets"));
                    });
            }
            let baseYml = docsDestination + '/powerpoint/powerpoint.yml';
            fsx.writeFileSync(baseYml, fsx.readFileSync(baseYml).toString()
                .replace(/\/office\/dev\/add-ins\/reference\/javascript-api-for-office/g, "/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets"));
        });

    console.log(`Fixing top href`);
    fsx.readdirSync(docsDestination)
        .filter(filename => filename.indexOf(".yml") < 0)
        .forEach(filename => {
            let subfolder = docsDestination + '/' + filename;
            fsx.readdirSync(subfolder)
                .filter(subfilename => subfilename.indexOf("toc") >= 0)
                .forEach(subfilename => {
                    fsx.writeFileSync(subfolder + '/' + subfilename, fsx.readFileSync(subfolder + '/' + subfilename).toString().replace("~/docs-ref-autogen/overview/office.md", "overview.md"));
                });
        });


    console.log(`Moving common TOC to its own folder`);
    fsx.copySync(docsDestination + "/office/toc.yml", docsDestination +  "/common/toc.yml");

    // remove to prevent build errors
    fsx.removeSync(docsDestination + "/office/overview.md");
    fsx.removeSync(docsDestination + "/office/toc.yml");

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

function scrubAndWriteToc(versionFolder: string, commonToc?: Toc, hostName?: string, versionNumber?: number): Toc {
    const tocPath = versionFolder + "/toc.yml";
    let latestToc;
    if (!commonToc) {
        latestToc = fixCommonToc(tocPath);
    } else {
        latestToc = fixToc(tocPath, commonToc, hostName, versionNumber);
    }

    fsx.writeFileSync(tocPath, jsyaml.safeDump(latestToc));
    return latestToc;
}

function fixToc(tocPath: string, commonToc: Toc, hostName: string, versionNumber: number): Toc {
    console.log(`Updating the structure of the TOC file: ${tocPath}`);

    let origToc = (jsyaml.safeLoad(fsx.readFileSync(tocPath).toString()) as Toc);
    let newToc = <Toc>{};
    let membersToMove = <IMembers>{};

    newToc.items = [{
        "name": "API reference",
        "items": [] as any
    }];
    newToc.items[0].items = [{
        "name": "API reference overview",
        "href": "../overview/overview.md"
    }] as any;

    let generalFilter: string[] = ["Interfaces"];
    let enumFilter: string[];
    if (hostName === "outlook") {
        enumFilter = ["MailboxEnums"];
        generalFilter = generalFilter.concat(enumFilter).concat(['Appointment', 'AppointmentForm', 'ItemCompose', 'ItemRead', 'Message']);
    } else {
        enumFilter = generateEnumList(fsx.readFileSync(`../api-extractor-inputs-${hostName}/${hostName}.d.ts`).toString());
        generalFilter = generalFilter.concat(enumFilter);
    }


    // create custom folders
    let excelIconSetFilter : string [] = ["FiveArrowsGraySet", "FiveArrowsSet", "FiveBoxesSet", "FiveQuartersSet", "FiveRatingSet", "FourArrowsGraySet", "FourArrowsSet", "FourRatingSet", "FourRedToBlackSet", "FourTrafficLightsSet", "IconCollections", "ThreeArrowsGraySet", "ThreeArrowsSet", "ThreeFlagsSet",  "ThreeSignsSet", "ThreeStarsSet",  "ThreeSymbols2Set", "ThreeSymbolsSet", "ThreeTrafficLights1Set", "ThreeTrafficLights2Set", "ThreeTrianglesSet"];
    let customFunctionsRoot = {"name": "Custom Functions", "uid": "", "items": [] as any};

    // create filter lists for types we shouldn't expose
    if (hostName === "excel") {
        generalFilter = generalFilter.concat(excelIconSetFilter);
    }

    origToc.items.forEach((rootItem, rootIndex) => {
        rootItem.items.forEach((packageItem, packageIndex) => {
            // fix host capitalization
            let packageName;
            if (packageItem.name === 'onenote') {
                packageName = 'OneNote';
            } else if (packageItem.name === 'powerpoint') {
                packageName = 'PowerPoint';
            } else {
                packageName = (packageItem.name.substr(0, 1).toUpperCase() + packageItem.name.substr(1)).replace(/\-/g, ' ');
            }

            // get items in the namespace for the new TOC
            packageItem.items.forEach((namespaceItem, namespaceIndex) => {
                membersToMove.items = namespaceItem.items;
            });

            if (packageName.toLocaleLowerCase().includes('custom functions runtime')) {
                customFunctionsRoot.items.push({
                    "name": packageName,
                    "uid": packageItem.uid,
                    "items":  membersToMove.items as any
                });
            } else {
                let primaryList = [] as any;
                if (packageName.toLocaleLowerCase().includes("excel")) {
                    // Excel has more subfolders and needs special casing to order them correctly.
                    let enumList = membersToMove.items.filter(item => {
                        return enumFilter.indexOf(item.name) >= 0;
                    });
                    let iconSetList = membersToMove.items.filter(item => {
                        return excelIconSetFilter.indexOf(item.name) >= 0;
                    });
                    primaryList = membersToMove.items.filter(item => {
                        return generalFilter.indexOf(item.name) < 0;
                    });

                    let excelEnumRoot = {"name": "Enums", "uid": "", "items": enumList};
                    let excelIconSetRoot = {"name": "Icon Sets", "uid": "", "items": iconSetList};
                    primaryList.unshift(excelIconSetRoot);
                    primaryList.unshift(excelEnumRoot);
                    if (versionNumber >= OLDEST_EXCEL_RELEASE_WITH_CUSTOM_FUNCTIONS) {
                        primaryList.unshift(customFunctionsRoot);
                    }
                } else if (membersToMove.items) {
                    let enumList = membersToMove.items.filter(item => {
                        return enumFilter.indexOf(item.name) >= 0;
                    });
                    primaryList = membersToMove.items.filter(item => {
                        return generalFilter.indexOf(item.name) < 0;
                    });

                    if (enumList) {
                        let enumRoot;
                        if (hostName === "outlook") {
                            enumRoot = membersToMove.items.find(item => {
                                return item.name === "MailboxEnums";
                            });
                        } else {
                            enumRoot = {"name": "Enums", "uid": "", "items": enumList};
                        }
                        primaryList.unshift(enumRoot);
                    }
                }

                newToc.items[0].items.push({
                    "name": packageName,
                    "uid": packageItem.uid,
                    "items": primaryList
                });
            }
        });
    });

    // Append the Common API TOC.
    if (hostName !== "visio") {
        newToc.items[0].items.push((commonToc.items[0] as any).items[1]);
    }

    return newToc;
}

function fixCommonToc(tocPath: string): Toc {
    console.log(`\nUpdating the structure of the Common TOC file: ${tocPath}`);

    let origToc = (jsyaml.safeLoad(fsx.readFileSync(tocPath).toString()) as Toc);
    let newToc = <Toc>{};

    newToc.items = [{
        "name": "API reference",
        "items": [] as any
    }];

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
            } else if (packageItem.name === 'office-runtime') {
                newToc.items[0].items[0].items.push({
                    "name": 'OfficeRuntime',
                    "uid": "office-runtime!",
                    "items": packageItem.items[0].items
                });
            }
        });
    });


    // add API reference overview to Common API
    newToc.items[0].items.unshift({
        "name": "API reference overview",
        "href": "../overview/overview.md"
    } as any);

    return newToc;
}
