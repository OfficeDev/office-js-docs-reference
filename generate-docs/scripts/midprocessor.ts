import * as fsx from 'fs-extra';
import * as path from "path";

tryCatch(async () => {
    const docsSource = path.resolve("../json");
    const docsDestination = path.resolve("../versioned-json");
    const snippets = path.resolve("../json/snippets.yaml");

    fsx.readdirSync(docsSource)
        .forEach(filename => {
            let subfolderName = docsDestination + '/' + filename.substring(0, filename.indexOf("."));
            fsx.copySync(
                docsSource + '/' + filename,
                subfolderName + '/' + filename
            );
            fsx.copySync(snippets, subfolderName + "/snippets.yaml");
        });
});

async function tryCatch(call: () => Promise<void>) {
    try {
        await call();
    } catch (e) {
        console.error(e);
        process.exit(1);
    }
}
