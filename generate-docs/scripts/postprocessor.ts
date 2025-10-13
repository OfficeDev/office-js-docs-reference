#!/usr/bin/env node --harmony

import { generateEnumList } from './util';
import * as fsx from 'fs-extra';
import * as jsyaml from "js-yaml";
import * as path from "path";

// Configuration constants
const OLDEST_EXCEL_RELEASE_WITH_CUSTOM_FUNCTIONS = 9;

const HOST_VERSION_MAP = [
    { host: "excel", versions: 21 }, // not including online
    { host: "onenote", versions: 1 },
    { host: "outlook", versions: 16 },
    { host: "powerpoint", versions: 10 },
    { host: "visio", versions: 1 },
    { host: "word", versions: 10 } // not including online or desktop
];

const EXCEL_ICON_SET_FILTER = [
    "FiveArrowsGraySet", "FiveArrowsSet", "FiveBoxesSet", "FiveQuartersSet", "FiveRatingSet",
    "FourArrowsGraySet", "FourArrowsSet", "FourRatingSet", "FourRedToBlackSet", "FourTrafficLightsSet",
    "IconCollections", "ThreeArrowsGraySet", "ThreeArrowsSet", "ThreeFlagsSet", "ThreeSignsSet",
    "ThreeStarsSet", "ThreeSymbols2Set", "ThreeSymbolsSet", "ThreeTrafficLights1Set",
    "ThreeTrafficLights2Set", "ThreeTrianglesSet"
];

const OUTLOOK_FILTER_ITEMS = ['Appointment', 'AppointmentForm', 'ItemCompose', 'ItemRead', 'Message'];

const NAMESPACE_REPLACEMENTS = {
    outlook: [
        { from: /CommonAPI/g, to: "Office" }
    ],
    office: [
        { from: /Outlook\.Mailbox/g, to: "Office.Mailbox" },
        { from: /Outlook\.RoamingSettings/g, to: "Office.RoamingSettings" },
        { from: /Outlook\.SensitivityLabelsCatalog/g, to: "Office.SensitivityLabelsCatalog" }
    ],
    customFunctions: [
        { from: /\/office\/dev\/add-ins\/reference\/javascript-api-for-office/g, to: "/javascript/api/requirement-sets/excel/custom-functions-requirement-sets" },
        { from: /\/office\/dev\/add-ins\/reference\/overview\/visio-javascript-reference-overview/g, to: "/javascript/api/requirement-sets/excel/custom-functions-requirement-sets" }
    ]
};

const SPECIAL_EXCEL_VERSIONS = [
    { folder: "excel_desktop_1_1", version: 20.5 },
];

const SPECIAL_WORD_VERSIONS = [
    { folder: "word_desktop_1_3", version: 9.10 },
    { folder: "word_desktop_1_2", version: 9.5 },
    { folder: "word_desktop_1_1", version: 8.5 },
    { folder: "word_1_5_hidden_document", version: 5.5 },
    { folder: "word_1_4_hidden_document", version: 4.5 },
    { folder: "word_1_3_hidden_document", version: 3.5 }
];

// File cleanup patterns
const YML_CLEANUP_PATTERNS = [
    { pattern: /^\s*example: \[\]\s*$/gm, replacement: "" },
    { pattern: /description: \\\*[\r\n]/gm, replacement: "description: ''" },
    { pattern: /\\\*/gm, replacement: "*" }
];

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

interface ApiFieldYaml {
    name: string;
    uid: string;
    package: string;
    summary: string;
    remarks?: string;
}

interface ApiPropertyYaml {
    name: string;
    uid: string;
    package: string;
    fullName: string;
    summary: string;
    remarks?: string;
    isPreview: boolean;
    isDeprecated: boolean;
    syntax: {
        content: string;
        return: {
            type: string;
            description?: string;
        }
    }
}

interface ApiMethodYaml {
    name: string;
    uid: string;
    package: string;
    fullName: string;
    summary: string;
    remarks?: string;
    isPreview: boolean;
    isDeprecated: boolean;
    syntax: {
        content: string;
        parameters?: {
            id: string;
            description: string;
            type: string;
        }[];
        return: {
            type: string;
            description: string;
        };
    };
}

interface ApiYaml {
    name: string;
    uid: string;
    package: string;
    fullName: string;
    summary: string;
    remarks: string;
    isPreview: boolean;
    isDeprecated: boolean;
    type: string;
    fields?: ApiFieldYaml[];
    properties?: ApiPropertyYaml[];
    methods?: ApiMethodYaml[];
    syntax?: string;
}

const docsSource = path.resolve("../yaml");
const docsDestination = path.resolve("../../docs/docs-ref-autogen");
const tocTemplateLocation = path.resolve("../../docs");

// Utility functions
function processFilesInDirectory(
    directory: string,
    filter: (filename: string) => boolean,
    processor: (filePath: string, content: string) => string
): void {
    if (!fsx.existsSync(directory)) return;
    
    fsx.readdirSync(directory)
        .filter(filter)
        .forEach(filename => {
            const filePath = path.join(directory, filename);
            const content = fsx.readFileSync(filePath, "utf8");
            const processedContent = processor(filePath, content);
            fsx.writeFileSync(filePath, processedContent);
        });
}

function applyNamespaceReplacements(content: string, replacements: Array<{ from: RegExp; to: string }>): string {
    return replacements.reduce((acc, { from, to }) => acc.replace(from, to), content);
}

function getHostNameFromFilename(filename: string): string {
    return filename.substring(0, filename.indexOf("_") < 0 ? filename.length : filename.indexOf("_"));
}

function capitalizeHostName(name: string): string {
    if (name === 'onenote') return 'OneNote';
    if (name === 'powerpoint') return 'PowerPoint';
    return name.charAt(0).toUpperCase() + name.slice(1).replace(/\-/g, ' ');
}

function createTocNode(name: string, uid?: string, items?: any[]): any {
    return { name, uid: uid || "", items: items || [] };
}

// Main processing functions
function cleanupOldDocs(): void {
    console.log(`Deleting old docs at: ${docsDestination}`);
    fsx.readdirSync(docsDestination)
        .filter(filename => filename !== "overview" && filename !== "images")
        .forEach(filename => fsx.removeSync(path.join(docsDestination, filename)));
}

function loadAndPrepareGlobalToc(): Toc {
    console.log(`Loading global TOC template`);
    let globalTocString = fsx.readFileSync(path.join(tocTemplateLocation, "toc.yml")).toString();
    globalTocString = globalTocString.replace(/href:\s*(.*)\.md/g, "href: ../../$1.md");
    return jsyaml.load(globalTocString) as Toc;
}

function copyDocsOutput(): void {
    console.log(`Copying docs output files to: ${docsDestination}`);
    fsx.readdirSync(docsSource).forEach(filename => {
        fsx.copySync(
            path.join(docsSource, filename),
            path.join(docsDestination, filename)
        );
    });
}

function processHostVersions(globalToc: Toc, tocWithPreviewCommon: Toc, tocWithReleaseCommon: Toc): void {
    HOST_VERSION_MAP.forEach(category => {
        const baseToc = category.host === "visio" ? globalToc : tocWithPreviewCommon;
        const versionToc = category.host === "visio" ? globalToc : tocWithReleaseCommon;

        if (category.versions > 1) {
            scrubAndWriteToc(path.join(docsDestination, category.host), baseToc, category.host, category.versions);
            for (let i = 1; i < category.versions; i++) {
                scrubAndWriteToc(path.join(docsDestination, `${category.host}_1_${i}`), versionToc, category.host, i);
            }
        } else {
            scrubAndWriteToc(path.join(docsDestination, category.host), versionToc, category.host, category.versions);
        }
    });
}

function processSpecialCases(tocWithReleaseCommon: Toc): void {
    // Special cases for Excel and Word Online
    scrubAndWriteToc(path.join(docsDestination, "excel_online"), tocWithReleaseCommon, "excel", 99);
    scrubAndWriteToc(path.join(docsDestination, "word_online"), tocWithReleaseCommon, "word", 99);

    // Special cases for Word Desktop versions
    SPECIAL_WORD_VERSIONS.forEach(({ folder, version }) => {
        scrubAndWriteToc(path.join(docsDestination, folder), tocWithReleaseCommon, "word", version);
    });
}

function processNamespaceReplacements(): void {
    console.log(`Namespace pass on Outlook docs`);
    fsx.readdirSync(docsDestination)
        .filter(filename => filename.includes("outlook") && !filename.includes(".yml"))
        .forEach(filename => {
            const subfolder = path.join(docsDestination, filename, "outlook");
            if (fsx.existsSync(subfolder)) {
                processFilesInDirectory(
                    subfolder,
                    () => true,
                    (_, content) => applyNamespaceReplacements(content, NAMESPACE_REPLACEMENTS.outlook)
                );
            }
        });

    console.log(`Namespace pass on Office docs`);
    const officeFolders = [
        path.join(docsDestination, "office", "office"),
        path.join(docsDestination, "office_release", "office")
    ];
    
    officeFolders.forEach(officeFolder => {
        console.log(officeFolder);
        if (fsx.existsSync(officeFolder)) {
            processFilesInDirectory(
                officeFolder,
                () => true,
                (_, content) => applyNamespaceReplacements(content, NAMESPACE_REPLACEMENTS.office)
            );
        }
    });
}

function processCustomFunctionsLinks(): void {
    console.log(`Custom Functions API requirement set link pass`);
    fsx.readdirSync(docsDestination)
        .filter(filename => filename.includes("excel") && !filename.includes(".yml"))
        .forEach(filename => {
            const subfolder = path.join(docsDestination, filename, "custom-functions-runtime");
            if (fsx.existsSync(subfolder)) {
                processFilesInDirectory(
                    subfolder,
                    () => true,
                    (_, content) => applyNamespaceReplacements(content, NAMESPACE_REPLACEMENTS.customFunctions)
                );
            }
        });
}

function processYamlFiles(): void {
    console.log(`Adjust YAML files - HREF and type alias expansion.`);
    fsx.readdirSync(docsDestination)
        .filter(filename => !filename.includes(".yml"))
        .forEach(filename => {
            const subfolder = path.join(docsDestination, filename);
            const hostName = getHostNameFromFilename(filename);
            
            if (fsx.existsSync(subfolder)) {
                fsx.readdirSync(subfolder).forEach(subfilename => {
                    const subfilePath = path.join(subfolder, subfilename);
                    
                    if (subfilename.includes("toc")) {
                        // Update overview HREF
                        const tocContent = fsx.readFileSync(subfilePath).toString()
                            .replace("~/docs-ref-autogen/overview/office.md", "overview.md");
                        fsx.writeFileSync(subfilePath, tocContent);
                    } else if (!subfilename.includes(".") && fsx.lstatSync(subfilePath).isDirectory()) {
                        // Package folder
                        processFilesInDirectory(
                            subfilePath,
                            fileName => fileName.includes(".yml"),
                            (_, ymlContent) => cleanUpYmlFile(ymlContent, hostName)
                        );
                    } else if (subfilename.includes(".yml")) {
                        const ymlContent = fsx.readFileSync(subfilePath, "utf8");
                        fsx.writeFileSync(subfilePath, cleanUpYmlFile(ymlContent, hostName));
                    }
                });
            }
        });
}

function moveCommonTocs(): void {
    console.log(`Moving common TOC to its own folder`);
    fsx.copySync(
        path.join(docsDestination, "office", "toc.yml"),
        path.join(docsDestination, "common_preview", "toc.yml")
    );
    fsx.copySync(
        path.join(docsDestination, "office_release", "toc.yml"),
        path.join(docsDestination, "common", "toc.yml")
    );
}

function cleanupTemporaryFiles(): void {
    // Remove files to prevent build errors
    const filesToRemove = [
        path.join(docsDestination, "office", "overview.md"),
        path.join(docsDestination, "office", "toc.yml"),
        path.join(docsDestination, "office_release", "toc.yml"),
        path.join(docsDestination, "office-runtime", "toc.yml")
    ];
    
    filesToRemove.forEach(file => fsx.removeSync(file));
}

tryCatch(async () => {
    console.log(`\nStarting postprocessor script...`);

    // Step 1: Clean up old documentation
    cleanupOldDocs();

    // Step 2: Load and prepare global TOC
    const globalToc = loadAndPrepareGlobalToc();

    // Step 3: Copy documentation output
    copyDocsOutput();

    // Step 4: Fix all the individual TOC files
    (globalToc.items[0].items[0] as ApplicationTocNode).href = "../overview/overview.md"; // Stay within a moniker
    const tocWithPreviewCommon = scrubAndWriteToc(path.join(docsDestination, "office"), globalToc);
    const tocWithReleaseCommon = scrubAndWriteToc(path.join(docsDestination, "office_release"), globalToc);

    // Step 5: Process host versions
    processHostVersions(globalToc, tocWithPreviewCommon, tocWithReleaseCommon);

    // Step 6: Process special cases
    processSpecialCases(tocWithReleaseCommon);

    // Step 7: Process namespace replacements
    processNamespaceReplacements();

    // Step 8: Process custom functions links
    processCustomFunctionsLinks();

    // Step 9: Process YAML files
    processYamlFiles();

    // Step 10: Move common TOCs and cleanup
    moveCommonTocs();
    cleanupTemporaryFiles();

    console.log(`\nPostprocessor script complete\n`);
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
    const tocPath = path.join(versionFolder, "toc.yml");
    let latestToc;
    if (!hostName) {
        latestToc = fixCommonToc(tocPath, globalToc);
    } else {
        latestToc = fixToc(tocPath, globalToc, hostName, versionNumber);
    }

    fsx.writeFileSync(tocPath, jsyaml.dump(latestToc));
    return latestToc;
}

function fixToc(tocPath: string, globalToc: Toc, hostName: string, versionNumber: number): Toc {
    console.log(`Updating the structure of the TOC file: ${tocPath}`);

    const origToc = jsyaml.load(fsx.readFileSync(tocPath).toString()) as Toc;
    let newTocNode = <ApplicationTocNode>{};
    let membersToMove = <IMembers>{};

    let generalFilter: string[] = ["Interfaces"];

    // Create custom folders and filters
    const customFunctionsRoot = createTocNode("Custom Functions", "", []);

    // Create filter lists for types we shouldn't expose
    if (hostName === "excel") {
        generalFilter = generalFilter.concat(EXCEL_ICON_SET_FILTER);
    } else if (hostName === "outlook") {
        generalFilter = generalFilter.concat(OUTLOOK_FILTER_ITEMS);
    }

    origToc.items.forEach((rootItem) => {
        rootItem.items.forEach((packageItem: ApplicationTocNode) => {
            // Fix host capitalization
            const packageName = capitalizeHostName(packageItem.name);

            // Get items in the namespace for the new TOC
            membersToMove.items = packageItem.items;

            if (packageName.toLowerCase().includes('custom functions runtime')) {
                customFunctionsRoot.items.push(createTocNode(packageName, packageItem.uid, membersToMove.items as any));
            } else {
                let primaryList = [] as any;
                if (membersToMove.items) {
                    const enumList = membersToMove.items.filter(item => item.uid.includes("enum"));
                    
                    primaryList = membersToMove.items.filter(item => {
                        return generalFilter.indexOf(item.name) < 0 && 
                               !item.uid.includes(".Interfaces.") && 
                               !item.uid.includes("enum");
                    });

                    if (enumList.length > 0) {
                        const enumRootName = packageName.toLowerCase().includes("outlook") ? "MailboxEnums" : "Enums";
                        const enumRoot = createTocNode(enumRootName, "", enumList);
                        
                        if (packageName.toLowerCase().includes("excel")) {
                            // Excel has subfolders for icon sets and custom functions
                            const iconSetList = membersToMove.items.filter(item => 
                                EXCEL_ICON_SET_FILTER.includes(item.name)
                            );

                            if (iconSetList.length > 0) {
                                const excelIconSetRoot = createTocNode("Icon Sets", "", iconSetList);
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
                    primaryList.forEach((namespaceItem) => {
                        if (namespaceItem.uid) {
                            const regex = /\w+\.(\w+\.\w+)/g;
                            const matchResults = regex.exec(namespaceItem.uid);
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

    return newToc;
}

function fixCommonToc(tocPath: string, globalToc: Toc): Toc {
    console.log(`\nUpdating the structure of the Common TOC file: ${tocPath}`);

    const origToc = jsyaml.load(fsx.readFileSync(tocPath).toString()) as Toc;
    const runtimeTocPath = path.resolve("../../docs/docs-ref-autogen/office-runtime/toc.yml");
    const runtimeToc = jsyaml.load(fsx.readFileSync(runtimeTocPath).toString()) as Toc;
    
    origToc.items[0].items = origToc.items[0].items.concat(runtimeToc.items[0].items);
    let membersToMove = <IMembers>{};

    // Create roots for items we want to reorder
    const newTocNode = {
        name: 'Common APIs',
        uid: "office!",
        items: [] as any
    };

    // Create folders for common (shared) API subcategories
    const officeTypesPath = path.resolve("../api-extractor-inputs-office/office.d.ts");
    const runtimeTypesPath = path.resolve("../api-extractor-inputs-office-runtime/office-runtime.d.ts");
    
    let sharedEnumFilter = generateEnumList(fsx.readFileSync(officeTypesPath).toString());
    sharedEnumFilter = sharedEnumFilter.concat(generateEnumList(fsx.readFileSync(runtimeTypesPath).toString()));

    // Process 'office' (Common "Shared" API) package
    origToc.items.forEach((rootItem) => {
        rootItem.items.forEach((packageItem: ApplicationTocNode) => {
            membersToMove.items = packageItem.items;
            
            if (packageItem.name.toLowerCase() === 'office') {
                membersToMove.items.forEach((namespaceItem) => {
                    // Scan UID for namespace to add to name
                    if (namespaceItem.uid) {
                        const regex = /\w+\.(\w+\.\w+)/g;
                        const matchResults = regex.exec(namespaceItem.uid);
                        if (matchResults) {
                            namespaceItem.name = matchResults[1];
                        }
                    }
                });

                const enumList = membersToMove.items.filter(item => 
                    sharedEnumFilter.includes(item.name)
                );
                const officeExtensionList = membersToMove.items.filter(item => 
                    item.uid.includes("office!OfficeExtension.")
                );
                const primaryList = membersToMove.items.filter(item => 
                    !sharedEnumFilter.includes(item.name) && !item.uid.includes("office!OfficeExtension.")
                );

                const sharedEnumRoot = createTocNode("Enums", "", enumList);
                primaryList.unshift(sharedEnumRoot);
                
                newTocNode.items.push(createTocNode('Office', packageItem.uid, primaryList));
                newTocNode.items.push(createTocNode('OfficeExtension', "", officeExtensionList));
                
            } else if (packageItem.name === 'office-runtime') {
                newTocNode.items.push(createTocNode('OfficeRuntime', packageItem.uid, packageItem.items));
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

    return newToc;
}

function cleanUpYmlFile(ymlFile: string, hostName: string): string {
    const schemaComment = ymlFile.substring(0, ymlFile.indexOf("\n") + 1);
    const apiYaml: ApiYaml = jsyaml.load(ymlFile) as ApiYaml;

    // Add links for type aliases.
    if (apiYaml.uid.endsWith(":type") && !apiYaml.uid.includes("Office")) {
        let remarks = `\n\nLearn more about the types in this type alias through the following links. \n\n`;
        const matches = apiYaml.syntax.substring(apiYaml.syntax.indexOf('=')).match(/[\w]+/g);
        
        if (matches) {
            matches.forEach((match, matchIndex) => {
                remarks += `[${capitalizeFirstLetter(hostName)}.${match}](/javascript/api/${hostName}/${hostName}.${match.toLowerCase()})`;
                if (matchIndex < matches.length - 1) {
                    remarks += ", ";
                }
            });
        }

        const exampleIndex = apiYaml.remarks.indexOf("#### Examples");
        if (exampleIndex > 0) {
            apiYaml.remarks = `${apiYaml.remarks.substring(0, exampleIndex)}${remarks}\n\n${apiYaml.remarks.substring(exampleIndex)}`;
        } else {
            apiYaml.remarks += remarks;
        }
    }
    
    let cleanYml = schemaComment + jsyaml.dump(apiYaml);
    
    // Apply cleanup patterns
    return YML_CLEANUP_PATTERNS.reduce((content, { pattern, replacement }) => 
        content.replace(pattern, replacement), cleanYml);
}

function capitalizeFirstLetter(str: string): string {
    if (!str) {
        return str;
    }
    return str.charAt(0).toUpperCase() + str.slice(1);
}
    