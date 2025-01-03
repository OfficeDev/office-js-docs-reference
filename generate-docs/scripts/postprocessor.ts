#!/usr/bin/env node --harmony

import { generateEnumList } from './util';
import * as fsx from 'fs-extra';
import * as jsyaml from "js-yaml";
import * as path from "path";

const OLDEST_EXCEL_RELEASE_WITH_CUSTOM_FUNCTIONS = 9;

interface Toc {
    items: [{
        name: string,
        href?: string,
        items?: (ApplicationTocNode | ManifestItem)[]
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

interface ManifestItem {
    name: string,
    href?: string,
    items: [
        {
            name: string,
            href: string
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

const docsSource = path.resolve("../yaml");
const docsDestination = path.resolve("../../docs/docs-ref-autogen");
const manifestRefPath = path.resolve("../../docs/manifest");
const requirementSetRefPath = path.resolve("../../docs/requirement-sets");

tryCatch(async () => {
    console.log("\nStarting postprocessor script...");

    console.log(`Deleting old docs at: ${docsDestination}`);
    // delete everything except the 'overview' folder from the /docs folder
    fsx.readdirSync(docsDestination)
        .filter(filename => filename !== "overview" && filename !== "images")
        .forEach(filename => fsx.removeSync(docsDestination + '/' + filename));

    console.log(`Creating global TOC`);
    let globalToc = <Toc>{items: [{name: "API reference"}]};
    globalToc.items[0].items = [{"name": "API reference overview", "href": "/javascript/api/overview"},
                                {"name": "Excel", "href": "/javascript/api/excel"},
                                {"name": "OneNote", "href": "/javascript/api/onenote"},
                                {"name": "Outlook", "href": "/javascript/api/outlook"},
                                {"name": "PowerPoint", "href": "/javascript/api/powerpoint"},
                                {"name": "Visio", "href": "/javascript/api/visio"},
                                {"name": "Word", "href": "/javascript/api/word"},
                                {"name": "Common APIs", "href": "/javascript/api/office"}] as any;

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
    (globalToc.items[0].items[0] as ApplicationTocNode).href = "../overview/overview.md"; // Stay within a moniker
    const tocWithPreviewCommon = scrubAndWriteToc(docsDestination + "/office", globalToc);
    const tocWithReleaseCommon = scrubAndWriteToc(docsDestination + "/office_release", globalToc);
    const hostVersionMap = [{host: "excel", versions: 18}, /*not including online*/
                            {host: "onenote", versions: 1},
                            {host: "outlook", versions: 15},
                            {host: "powerpoint", versions: 8},
                            {host: "visio", versions: 1},
                            {host: "word", versions: 10}]; /* not including online or desktop*/

    hostVersionMap.forEach(category => {
        if (category.versions > 1) {
            scrubAndWriteToc(path.resolve(`${docsDestination}/${category.host}`), category.host === "visio" ? globalToc : tocWithPreviewCommon, category.host, category.versions);
            for (let i = 1; i < category.versions; i++) {
                scrubAndWriteToc(path.resolve(`${docsDestination}/${category.host}_1_${i}`), category.host === "visio" ? globalToc : tocWithReleaseCommon, category.host, i);
            }
        } else {
            // This assumes the single version of the application's docs is not a preview version.
            scrubAndWriteToc(path.resolve(`${docsDestination}/${category.host}`), category.host === "visio" ? globalToc : tocWithReleaseCommon, category.host, category.versions);
        }
    });

    // Special case for ExcelApi Online
    scrubAndWriteToc(path.resolve(`${docsDestination}/excel_online`), tocWithReleaseCommon, "excel", 99);

    // Special case for WordApi Online
    scrubAndWriteToc(path.resolve(`${docsDestination}/word_online`), tocWithReleaseCommon, "word", 99);

    // Special case for WordApi Desktop
    scrubAndWriteToc(path.resolve(`${docsDestination}/word_desktop_1_1`), tocWithReleaseCommon, "word", 8.5);
    scrubAndWriteToc(path.resolve(`${docsDestination}/word_1_5_hidden_document`), tocWithReleaseCommon, "word", 5.5);
    scrubAndWriteToc(path.resolve(`${docsDestination}/word_1_4_hidden_document`), tocWithReleaseCommon, "word", 4.5);
    scrubAndWriteToc(path.resolve(`${docsDestination}/word_1_3_hidden_document`), tocWithReleaseCommon, "word", 3.5);

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
    const officeFolders: string[] = [docsDestination + "/office/office", docsDestination + "/office_release/office"];
    officeFolders.forEach((officeFolder) => {
    console.log(officeFolder);
        fsx.readdirSync(officeFolder)
            .forEach(filename => {
                fsx.writeFileSync(officeFolder + '/' + filename, fsx.readFileSync(officeFolder + '/' + filename).toString().replace(/Outlook\.Mailbox/g, "Office.Mailbox").replace(/Outlook\.RoamingSettings/g, "Office.RoamingSettings").replace(/Outlook\.SensitivityLabelsCatalog/g, "Office.SensitivityLabelsCatalog"));
            });
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
                                .replace(/\/office\/dev\/add-ins\/reference\/javascript-api-for-office/g, "/javascript/api/requirement-sets/excel/custom-functions-requirement-sets")
                                .replace(/\/office\/dev\/add-ins\/reference\/overview\/visio-javascript-reference-overview/g, "/javascript/api/requirement-sets/excel/custom-functions-requirement-sets"));
                    });
            }
        });

    console.log(`Fixing top href`);
    fsx.readdirSync(docsDestination)
        .filter(filename => filename.indexOf(".yml") < 0)
        .forEach(filename => {
            let subfolder = docsDestination + '/' + filename;
            fsx.readdirSync(subfolder).forEach(subfilename => {
                    if (subfilename.indexOf("toc") >= 0) {
                        // Update overview HREF.
                        fsx.writeFileSync(subfolder + '/' + subfilename, fsx.readFileSync(subfolder + '/' + subfilename).toString().replace("~/docs-ref-autogen/overview/office.md", "overview.md"));
                    } else if (subfilename.indexOf(".") < 0) {
                        let packageFolder = subfolder + '/' + subfilename;
                        fsx.readdirSync(packageFolder).filter(packageFileName => packageFileName.indexOf(".yml") > 0).forEach(packageFileName => {
                            
                            fsx.writeFileSync(packageFolder + '/' + packageFileName, fsx.readFileSync(packageFolder + '/' + packageFileName).toString()
                                .replace(/^\s*example: \[\]\s*$/gm, "") // Remove example field from yml as the OPS schema does not support it.
                                .replace(/description: \\\*[\r\n]/gm, "description: ''") // Remove descriptions that are just "\*".
                                .replace(/\\\*/gm, "*")); // Fix asterisk protection.
                        });
                    } else if (subfilename.indexOf(".yml") > 0) {
                        
                        fsx.writeFileSync(subfolder + '/' + subfilename, fsx.readFileSync(subfolder + '/' + subfilename).toString()
                            .replace(/^\s*example: \[\]\s*$/gm, "") // Remove example field from yml as the OPS schema does not support it.
                            .replace(/description: \\\*[\r\n]/g, "description:") // Remove descriptions that are just "\*".
                            .replace(/\\\*/gm, "*")); // Fix asterisk protection.
                    }
                });
        });

    console.log(`Moving common TOC to its own folder`);
    fsx.copySync(docsDestination + "/office/toc.yml", docsDestination +  "/common_preview/toc.yml");
    fsx.copySync(docsDestination + "/office_release/toc.yml", docsDestination +  "/common/toc.yml");

    // remove to prevent build errors
    fsx.removeSync(docsDestination + "/office/overview.md");
    fsx.removeSync(docsDestination + "/office/toc.yml");
    fsx.removeSync(docsDestination + "/office_release/toc.yml");
    fsx.removeSync(docsDestination + "/office-runtime/toc.yml");

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
        latestToc = fixToc(tocPath, globalToc, hostName, versionNumber, hostName === "visio");
    }

    fsx.writeFileSync(tocPath, jsyaml.safeDump(latestToc));
    return latestToc;
}

function fixToc(tocPath: string, globalToc: Toc, hostName: string, versionNumber: number, addNonApiRef: boolean): Toc {
    console.log(`Updating the structure of the TOC file: ${tocPath}`);

    let origToc = (jsyaml.safeLoad(fsx.readFileSync(tocPath).toString()) as Toc);
    let newTocNode = <ApplicationTocNode>{};
    let membersToMove = <IMembers>{};

    let generalFilter: string[] = ["Interfaces"];

    // create custom folders
    let excelIconSetFilter : string [] = ["FiveArrowsGraySet", "FiveArrowsSet", "FiveBoxesSet", "FiveQuartersSet", "FiveRatingSet", "FourArrowsGraySet", "FourArrowsSet", "FourRatingSet", "FourRedToBlackSet", "FourTrafficLightsSet", "IconCollections", "ThreeArrowsGraySet", "ThreeArrowsSet", "ThreeFlagsSet",  "ThreeSignsSet", "ThreeStarsSet",  "ThreeSymbols2Set", "ThreeSymbolsSet", "ThreeTrafficLights1Set", "ThreeTrafficLights2Set", "ThreeTrianglesSet"];
    let customFunctionsRoot = {"name": "Custom Functions", "uid": "", "items": [] as any};

    // create filter lists for types we shouldn't expose
    if (hostName === "excel") {
        generalFilter = generalFilter.concat(excelIconSetFilter);
    } else if (hostName === "outlook") {
        generalFilter = generalFilter.concat(['Appointment', 'AppointmentForm', 'ItemCompose', 'ItemRead', 'Message']);
    }

    origToc.items.forEach((rootItem, rootIndex) => {
        rootItem.items.forEach((packageItem: ApplicationTocNode, packageIndex) => {
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
                        return item.uid.indexOf("enum") >= 0;
                    });
                    primaryList = membersToMove.items.filter(item => {
                        // Remove previous chosen items and anything with the "Interfaces" namespace (those are Rich API duplicates for load/set).
                        return generalFilter.indexOf(item.name) < 0 && item.uid.indexOf(".Interfaces.") < 0 && item.uid.indexOf("enum") < 0;
                    });

                    if (enumList) {
                        const enumRootName = packageName.toLocaleLowerCase().includes("outlook") ? "MailboxEnums" : "Enums";
                        let enumRoot = {"name": enumRootName, "uid": "", "items": enumList};
                        if (packageName.toLocaleLowerCase().includes("excel")) {
                            // Excel has also has subfolders for icon sets and custom functions. They need to be correctly ordered.
                            let iconSetList = membersToMove.items.filter(item => {
                                return excelIconSetFilter.indexOf(item.name) >= 0;
                            });

                            if (iconSetList.length > 0) {
                                let excelIconSetRoot = {"name": "Icon Sets", "uid": "", "items": iconSetList};
                                primaryList.unshift(excelIconSetRoot);
                            }
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
                            let regex = /\w+\.(\w+\.\w+)/g;
                            let matchResults = regex.exec(namespaceItem.uid);
                            if (matchResults) {
                                namespaceItem.name = matchResults[1];
                            }
                        }
                    });
                }

                newTocNode = {
                    name: packageName,
                    uid: packageItem.uid,
                    items: primaryList
                };
            }
        });
    });

    const newToc = <Toc>{items: [] as any};
    globalToc.items.forEach((topLevel, topLevelIndex) => {
        newToc.items.push({name: topLevel.name, items: []});
        topLevel.items.forEach((applicationNode) => {
            if (applicationNode.name === newTocNode.name) {
                newToc.items[topLevelIndex].items.push(newTocNode);
            } else {
                newToc.items[topLevelIndex].items.push(applicationNode);
            }
        });
    });

    if (addNonApiRef) { 
        addNonRefTocInformation(newToc);
    }

    return newToc;
}

function fixCommonToc(tocPath: string, globalToc: Toc): Toc {
    console.log(`\nUpdating the structure of the Common TOC file: ${tocPath}`);

    let origToc = (jsyaml.safeLoad(fsx.readFileSync(tocPath).toString()) as Toc);
    let runtimeToc = (jsyaml.safeLoad(fsx.readFileSync(path.resolve("../../docs/docs-ref-autogen/office-runtime/toc.yml")).toString()) as Toc);
    origToc.items[0].items = origToc.items[0].items.concat(runtimeToc.items[0].items);
    let membersToMove = <IMembers>{};

    // Create roots for items we want to reorder.
    let newTocNode = {
        name: 'Common APIs',
        uid: "office!",
        items: [] as any
    };

    // create folders for common (shared) API subcategories
    let sharedEnumFilter = generateEnumList(fsx.readFileSync("../api-extractor-inputs-office/office.d.ts").toString());
    sharedEnumFilter.concat(generateEnumList(fsx.readFileSync("../api-extractor-inputs-office-runtime/office-runtime.d.ts").toString()));

    // process 'office' (Common "Shared" API) package
    origToc.items.forEach((rootItem, rootIndex) => {
        rootItem.items.forEach((packageItem: ApplicationTocNode, packageIndex) => {
            membersToMove.items = packageItem.items;
            if (packageItem.name.toLocaleLowerCase() === 'office') {
                membersToMove.items.forEach((namespaceItem, namespaceIndex) => {
                    // Scan UID for namespace to add to name.
                     if (namespaceItem.uid) {
                        let regex = /\w+\.(\w+\.\w+)/g;
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
    globalToc.items.forEach((topLevel, topLevelIndex) => {
        newToc.items.push({name: topLevel.name, items: []});
        topLevel.items.forEach((applicationNode) => {
            if (applicationNode.name === newTocNode.name) {
                newToc.items[topLevelIndex].items.push(newTocNode);
            } else {
                newToc.items[topLevelIndex].items.push(applicationNode);
            }
        });
    });

    addNonRefTocInformation(newToc);
    return newToc;
}

function addNonRefTocInformation(toc: Toc) {
    // Add manifest TOC
    let manifestYml = jsyaml.safeLoad(fsx.readFileSync(`${manifestRefPath}/toc.yml`).toString());
    toc.items.push(manifestYml[0]);

    // Add requirement sets TOC
    let requirementSetYml = jsyaml.safeLoad(fsx.readFileSync(`${requirementSetRefPath}/toc.yml`).toString());
    toc.items.push(requirementSetYml[0]);
}
