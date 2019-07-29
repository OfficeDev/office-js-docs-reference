import * as fsx from "fs-extra";
import { diff } from 'deep-diff';
import { APISet, parseDTS } from './dts-utilities';

tryCatch(async () => {
    //fsx.writeFileSync("./tool-inputs/dt-release.d.ts", await fetchAndThrowOnError("https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts", "text"));
    //fsx.writeFileSync("./tool-inputs/cdn-release.d.ts", await fetchAndThrowOnError("https://appsforoffice.officeapps.live.com/lib/1.1/hosted/office.d.ts", "text"));
    // read whole d.ts files
    const wholeReleaseDT = fsx.readFileSync("./tool-inputs/dt-release.d.ts").toString().replace(/\r\n/gm, "\n");
    //const wholePreviewDT = fsx.readFileSync("https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts").toString();
    const wholeReleaseCDN = fsx.readFileSync("./tool-inputs/cdn-release.d.ts").toString().replace(/\r\n/gm, "\n");
    //const wholePreviewCDN = fsx.readFileSync("https://appsforoffice.officeapps.live.com/lib/beta/hosted/office.d.ts").toString();


    const releaseDTAPI: APISet = parseDTS("index", wholeReleaseDT);
    //const previewDTAPI: APISet = parseDTS("PreviewDT", wholePreviewDT);
    const releaseCDNAPI: APISet = parseDTS("ReleaseDT", wholeReleaseCDN);
    //const previewCDNAPI: APISet = parseDTS("ReleaseDT", wholePreviewCDN);

    let diffedObject = diff(releaseDTAPI, releaseCDNAPI);
    console.log(diffedObject);
});

async function tryCatch(call: () => Promise<void>) {
    try {
        await call();
    } catch (e) {
        console.error(e);
        process.exit(1);
    }
}
