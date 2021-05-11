#!/usr/bin/env node --harmony

import { generateEnumList } from './util';
import * as fsx from 'fs-extra';
import * as jsyaml from "js-yaml";
import * as path from "path";

const OLDEST_EXCEL_RELEASE_WITH_CUSTOM_FUNCTIONS = 9;

class Toc {
    
    items: [{
        name: string,
        items: ApplicationTocNode[]
    }]
}

interface ApplicationTocNode {
    name: string,
    href?: string,
    uid?: string
    items: [
        {
            name: string,
            uid: string,
            items: [
                {
                    name: string,
                    uid?: string,
                    items: IMembers
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
                                {"name": "Excel", "href": "/javascript/api/excel"},
                                {"name": "OneNote", "href": "/javascript/api/onenote"},
                                {"name": "Outlook", "href": "/javascript/api/outlook"},
                                {"name": "PowerPoint", "href": "/javascript/api/powerpoint"},
                                {"name": "Visio", "href": "/javascript/api/visio"},
                                {"name": "Word", "href": "/javascript/api/word"},
                                {"name": "Common APIs", "href": "/javascript/api/office"}] as any;
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
    globalToc.items[0].items[0].href = "../overview/overview.md"; // Stay within a moniker
    const tocWithCommon = scrubAndWriteToc(docsDestination + "/office", globalToc);
    const hostVersionMap = [{host: "excel", versions: 13}, /*not including online*/
                            {host: "onenote", versions: 1},
                            {host: "outlook", versions: 11},
                            {host: "powerpoint", versions: 3},
                            {host: "visio", versions: 1},
                            {host: "word", versions: 4}];

    hostVersionMap.forEach(category => {
        let tocToUse = category.host === "visio" ? globalToc : tocWithCommon; // Visio doesn't have access to Common APIs.
        scrubAndWriteToc(path.resolve(`${docsDestination}/${category.host}`), tocToUse, category.host, category.versions);
        for (let i = 1; i < category.versions; i++) {
            scrubAndWriteToc(path.resolve(`${docsDestination}/${category.host}_1_${i}`), tocToUse, category.host, i);
        }
    });

    // Special case for ExcelApi Online
    scrubAndWriteToc(path.resolve(`${docsDestination}/excel_online`), tocWithCommon, "excel", 99);


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

function scrubAndWriteToc(versionFolder: string, globalToc: Toc, hostName?: string, versionNumber?: number): Toc {
    const tocPath = versionFolder + "/toc.yml";
    let latestToc;
    if (!hostName) {
        latestToc = fixCommonToc(tocPath, globalToc);
    } else {
        latestToc = fixToc(tocPath, globalToc, hostName, versionNumber);
    }

    fsx.writeFileSync(tocPath, jsyaml.safeDump(latestToc));
    return latestToc;
}

function fixToc(tocPath: string, globalToc: Toc, hostName: string, versionNumber: number): Toc {
    console.log(`Updating the structure of the TOC file: ${tocPath}`);

    let origToc = (jsyaml.safeLoad(fsx.readFileSync(tocPath).toString()) as Toc);
    let newTocNode = <ApplicationTocNode>{};
    let membersToMove = <IMembers>{};

    let generalFilter: string[] = ["Interfaces"];
    let enumFilter: string[] = generateEnumList(fsx.readFileSync(`../api-extractor-inputs-${hostName}/${hostName}.d.ts`).toString());
    generalFilter = generalFilter.concat(enumFilter);

    // create custom folders
    let excelIconSetFilter : string [] = ["FiveArrowsGraySet", "FiveArrowsSet", "FiveBoxesSet", "FiveQuartersSet", "FiveRatingSet", "FourArrowsGraySet", "FourArrowsSet", "FourRatingSet", "FourRedToBlackSet", "FourTrafficLightsSet", "IconCollections", "ThreeArrowsGraySet", "ThreeArrowsSet", "ThreeFlagsSet",  "ThreeSignsSet", "ThreeStarsSet",  "ThreeSymbols2Set", "ThreeSymbolsSet", "ThreeTrafficLights1Set", "ThreeTrafficLights2Set", "ThreeTrianglesSet"];
    let customFunctionsRoot = {"name": "Custom Functions", "uid": "", "items": [] as any};

    // create filter lists for types we shouldn't expose
    if (hostName === "excel") {
        generalFilter = generalFilter.concat(excelIconSetFilter);
    } else if (hostName === "outlook") {
        generalFilter = generalFilter.concat(enumFilter).concat(['Appointment', 'AppointmentForm', 'ItemCompose', 'ItemRead', 'Message']);
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
            membersToMove.items = packageItem.items;

            if (packageName.toLocaleLowerCase().includes('custom functions runtime')) {
                customFunctionsRoot.items.push({
                    "name": packageName,
                    "uid": packageItem.uid,
                    "items":  membersToMove.items as any
                });
            } else {
                let primaryList = [] as any;
                if (membersToMove.items) {
                    let enumList = membersToMove.items.filter(item => {
                        return enumFilter.indexOf(item.name) >= 0;
                    });
                    primaryList = membersToMove.items.filter(item => {
                        // Remove previous chosen items and anything with the "Interfaces" namespace (those are Rich API duplicates for load/set).
                        return generalFilter.indexOf(item.name) < 0 && item.uid.indexOf(".Interfaces.") < 0;
                    });

                    if (enumList) {
                        const enumRootName = packageName.toLocaleLowerCase().includes("outlook") ? "MailboxEnums" : "Enums";
                        let enumRoot = {"name": enumRootName, "uid": "", "items": enumList};
                        if (packageName.toLocaleLowerCase().includes("excel")) {
                            // Excel has also has subfolders for icon sets and custom functions. They need to be correctly ordered.
                            let iconSetList = membersToMove.items.filter(item => {
                                return excelIconSetFilter.indexOf(item.name) >= 0;
                            });
        
                            let excelIconSetRoot = {"name": "Icon Sets", "uid": "", "items": iconSetList};
                            primaryList.unshift(excelIconSetRoot);
                            primaryList.unshift(enumRoot);            
                            if (versionNumber >= OLDEST_EXCEL_RELEASE_WITH_CUSTOM_FUNCTIONS) {
                                primaryList.unshift(customFunctionsRoot);
                            }
                        } else {
                            primaryList.unshift(enumRoot);
                        }
                    }                    

                    // Address any nested namespaces
                    primaryList.forEach((namespaceItem, namespaceIndex) => {
                        // Scan UID for namespace to add to name.
                        if (namespaceItem.uid) {
                            let regex = /\w+\.(\w+\.\w+)/g
                            let matchResults = regex.exec(namespaceItem.uid);
                            if (matchResults) {
                                namespaceItem.name = matchResults[1];
                            }
                        }
                    });
                }

                newTocNode= {
                    name: packageName,
                    uid: packageItem.uid,
                    items: primaryList
                };
            }
        });
    });

    const newToc = <Toc>{items: [] as any};
    globalToc.items.forEach((topLevel, topLevelIndex) =>{
        newToc.items.push({name: topLevel.name, items: []});
        topLevel.items.forEach((applicationNode) =>{
            if (applicationNode.name === newTocNode.name) {
                newToc.items[topLevelIndex].items.push(newTocNode);
            } else {
                newToc.items[topLevelIndex].items.push(applicationNode);
            }
        });
    });

    return newToc;
}

function fixCommonToc(tocPath: string, globalToc: Toc): Toc {
    console.log(`\nUpdating the structure of the Common TOC file: ${tocPath}`);

    let origToc = (jsyaml.safeLoad(fsx.readFileSync(tocPath).toString()) as Toc);
    let membersToMove = <IMembers>{};

    // Create roots for items we want to reorder.
    let newTocNode = {
        name: 'Common APIs',
        uid: "office!",
        items: [] as any
    }

    // create folders for common (shared) API subcategories
    let sharedEnumFilter = generateEnumList(fsx.readFileSync("../api-extractor-inputs-office/office.d.ts").toString());

    // process 'office' (Common "Shared" API) package
    origToc.items.forEach((rootItem, rootIndex) => {
        rootItem.items.forEach((packageItem, packageIndex) => {
            membersToMove.items = packageItem.items;
            if (packageItem.name.toLocaleLowerCase() === 'office') {
                membersToMove.items.forEach((namespaceItem, namespaceIndex) => {                    
                    // Scan UID for namespace to add to name.
                     if (namespaceItem.uid) {
                        let regex = /\w+\.(\w+\.\w+)/g
                        let matchResults = regex.exec(namespaceItem.uid);
                        if (matchResults) {
                            namespaceItem.name = matchResults[1];
                        }
                    }
                });

                let enumList = membersToMove.items.filter(item => {
                    return sharedEnumFilter.indexOf(item.name) >= 0;
                });
                let officeExtensionList = membersToMove.items.filter(item => {
                    return item.uid.indexOf("office!OfficeExtension.") >= 0;
                });
                let primaryList = membersToMove.items.filter(item => {
                    return sharedEnumFilter.indexOf(item.name) < 0 && item.uid.indexOf("office!OfficeExtension.") < 0;
                });

                let sharedEnumRoot = {"name": "Enums", "uid": "", "items": enumList};
                primaryList.unshift(sharedEnumRoot);            
                newTocNode.items.push({
                    "name": 'Office',
                    "uid": packageItem.uid,
                    "items": primaryList
                });
                newTocNode.items.push({
                    "name": 'OfficeExtension',
                    "items": officeExtensionList
                });
            } else if (packageItem.name === 'office-runtime') {
                newTocNode.items.push({
                    "name": 'OfficeRuntime',
                    "uid": packageItem.uid,
                    "items": packageItem.items
                });
            }
        });
    });

    const newToc = <Toc>{items: [] as any};
    globalToc.items.forEach((topLevel, topLevelIndex) =>{
        newToc.items.push({name: topLevel.name, items: []});
        topLevel.items.forEach((applicationNode) =>{
            if (applicationNode.name === newTocNode.name) {
                newToc.items[topLevelIndex].items.push(newTocNode);
            } else {
                newToc.items[topLevelIndex].items.push(applicationNode);
            }
        });
    });

    return newToc;
}
