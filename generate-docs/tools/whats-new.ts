import * as fsx from "fs-extra";
import { APISet, parseDTS } from './dts-utilities';


if (process.argv.length !== 6 || process.argv.find((x: string) => {return x === "-?"})) {
    console.log("usage: node whats-new [new d.ts] [old d.ts] [output file name (minus extension)]");
    process.exit(0);
}

const hostName = process.argv[2];
const newDtsPath = process.argv[3];
const oldDtsPath = process.argv[4];
const outputPath = process.argv[5];

console.log(`What's New between ${newDtsPath} and ${oldDtsPath}?`);

tryCatch(async () => {
    // read whole files
    let wholeRelease = fsx.readFileSync(oldDtsPath).toString();
    let wholePreview = fsx.readFileSync(newDtsPath).toString();

    const releaseAPI: APISet = parseDTS("release", wholeRelease);
    const previewAPI: APISet = parseDTS("preview", wholePreview);
    const diffAPI: APISet = previewAPI.diff(releaseAPI);

    const relativePath: string = "javascript/api/" + hostName.toLowerCase() + "/" + hostName.toLowerCase() + ".";
    if (!fsx.existsSync(outputPath + ".md")) {
        fsx.createFileSync(outputPath + ".md");
    }

    fsx.writeFileSync(outputPath + ".md", diffAPI.getAsMarkdown(relativePath));
});

async function tryCatch(call: () => Promise<void>) {
    try {
        await call();
    } catch (e) {
        console.error(e);
        process.exit(1);
    }
}
