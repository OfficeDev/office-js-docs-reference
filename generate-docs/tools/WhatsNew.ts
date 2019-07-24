import { readFileSync } from "fs";
import { promptFromList } from '../scripts/simple-prompts';
import { fetchAndThrowOnError, DtsBuilder} from '../scripts/util';
import * as fsx from "fs-extra";
import * as ts from "typescript";
import { APISet, parseDTS } from './dts-utilities';

tryCatch(async () => {
    // Get file locations
    const officeJSUrl = await promptFromList({
        message: "Which d.ts file should be used as the RELEASE version?",
        choices: [
            { name: "DefinitelyTyped", value: "https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts" },
            { name: "Local file [generate-docs\\tools\\tool-inputs\\release.d.ts]", value: "" }
        ]
    });

    if (officeJSUrl.length > 0) {
        fsx.writeFileSync("./tool-inputs/release.d.ts", await fetchAndThrowOnError(officeJSUrl, "text"));
    }

    const officeJSPreviewUrl = await promptFromList({
        message: "Which d.ts file should be used as the PREVIEW version?",
        choices: [
            { name: "DefinitelyTyped", value: "https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts" },
            { name: "Local file [generate-docs\\tools\\tool-inputs\\preview.d.ts]", value: "" }
        ]
    });

    if (officeJSPreviewUrl.length > 0) {
        fsx.writeFileSync("./tool-inputs/preview.d.ts", await fetchAndThrowOnError(officeJSPreviewUrl, "text"));
    }

    // read whole files
    let wholeRelease = fsx.readFileSync("./tool-inputs/release.d.ts").toString();
    let wholePreview = fsx.readFileSync("./tool-inputs/preview.d.ts").toString();

    const hostName = await promptFromList({
        message: "Which host is being generated?",
        choices: [
            { name: "Excel", value: "Excel" },
            { name: "OneNote", value: "OneNote" },
            { name: "Outlook", value: "Mailbox" },
            { name: "Visio", value: "Visio" },
            { name: "Word", value: "Word" }
        ]
    });
    const releaseHostFileName: string = './tool-inputs/' + hostName + '-release.d.ts';
    const previewHostFileName: string = './tool-inputs/' + hostName + '-preview.d.ts';

    const dtsBuilder = new DtsBuilder();
    fsx.writeFileSync(
        './tool-inputs/' + hostName + '-release.d.ts',
        dtsBuilder.extractDtsSection(wholeRelease, "Begin " + hostName + " APIs", "End " + hostName + " APIs")
    );
    fsx.writeFileSync(
        './tool-inputs/' + hostName + '-preview.d.ts',
        dtsBuilder.extractDtsSection(wholePreview, "Begin " + hostName + " APIs", "End " + hostName + " APIs")
    );

    const releaseAPI: APISet = new APISet();
    const previewAPI: APISet = new APISet();

    const releaseFile: ts.SourceFile = ts.createSourceFile(
        "Release",
        readFileSync(releaseHostFileName).toString(),
        ts.ScriptTarget.ES2015,
        true);
    const previewFile: ts.SourceFile = ts.createSourceFile(
        "Preview",
        readFileSync(previewHostFileName).toString(),
        ts.ScriptTarget.ES2015,
        true);

    parseDTS(releaseFile, releaseAPI);
    parseDTS(previewFile, previewAPI);

    const diffAPI: APISet = previewAPI.diff(releaseAPI);

    const relativePath: string = "javascript/api/" + hostName.toLowerCase() + "/" + hostName.toLowerCase() + ".";
    fsx.writeFileSync("./tool-outputs/WhatsNew.d.ts", diffAPI.getAsDTS());
    fsx.writeFileSync("./tool-outputs/WhatsNew.md", diffAPI.getAsMarkdown(relativePath));
});

async function tryCatch(call: () => Promise<void>) {
    try {
        await call();
    } catch (e) {
        console.error(e);
        process.exit(1);
    }
}
