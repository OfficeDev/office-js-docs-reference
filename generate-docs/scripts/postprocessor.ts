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
                            name: string, //kb: need to add optional items property...for office namespaces
                            uid?: string
                        }
                    ]
                }
            ]
        }
    ]
}

tryCatch(async () => {

    console.log(`Updating the TOC file: ${path.resolve("../generate-docs/yaml/toc.yml")}`);

    let origToc = (jsyaml.safeLoad(fsx.readFileSync("../generate-docs/yaml/toc.yml").toString()) as IOrigToc);

    let newToc = <INewToc>{};

    let membersToMove = <IMembers>{};

    newToc.items = [{
        "name": origToc.items[0].name,
        "href": origToc.items[0].href
    }];
    newToc.items[0].items = [] as any;

    origToc.items.forEach((rootItem, rootIndex) => {
        rootItem.items.forEach((packageItem, packageIndex) => {
            if (packageItem.name !== 'office') {
                packageItem.items.forEach((namespaceItem, index) => {
                    membersToMove.items = namespaceItem.items;
                });
                newToc.items[0].items.push({
                    "name": packageItem.name.substr(0, 1).toUpperCase() + packageItem.name.substr(1),
                    "uid": packageItem.uid,
                    "items": membersToMove.items
                });
            }
            console.log(packageItem.name + '<br/>');
        });
    });

    origToc.items.forEach((rootItem, rootIndex) => {
        rootItem.items.forEach((packageItem, packageIndex) => {
            if (packageItem.name === 'office') {
                newToc.items[0].items.push({
                    "name": 'Shared API',
                    "uid": packageItem.uid,
                });

                packageItem.items.forEach((namespaceItem, namespaceIndex) => {
                    membersToMove.items = namespaceItem.items;
                    newToc.items[0].items[namespaceIndex].items.push({ 
                        "name": namespaceItem.name,
                        "items": membersToMove.items
                    });
                });
            }
            console.log(packageItem.name + '<br/>');
        });
    });

    // write file
    fsx.writeFileSync("zNewToc.yml", jsyaml.safeDump(newToc));

    console.log("Done!");

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

