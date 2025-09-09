"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
const fsx = require("fs-extra");
const dts_utilities_1 = require("./dts-utilities");
if (process.argv.length !== 6 || process.argv.find((x) => { return x === "-?"; })) {
    console.log("usage: node whats-new [new d.ts] [old d.ts] [output file name (minus extension)]");
    process.exit(0);
}
const hostName = process.argv[2];
const newDtsPath = process.argv[3];
const oldDtsPath = process.argv[4];
const outputPath = process.argv[5];
console.log(`What's New between ${newDtsPath} and ${oldDtsPath}?`);
tryCatch(() => __awaiter(void 0, void 0, void 0, function* () {
    // read whole files
    let wholeRelease = fsx.readFileSync(oldDtsPath).toString();
    let wholePreview = fsx.readFileSync(newDtsPath).toString();
    const releaseAPI = dts_utilities_1.parseDTS("release", wholeRelease);
    const previewAPI = dts_utilities_1.parseDTS("preview", wholePreview);
    const diffAPI = previewAPI.diff(releaseAPI);
    const relativePath = `javascript/api/${hostName.toLowerCase()}/${hostName.toLowerCase() === "outlook" ? "office" : hostName.toLowerCase()}.`;
    if (!fsx.existsSync(outputPath + ".md")) {
        fsx.createFileSync(outputPath + ".md");
    }
    fsx.writeFileSync(outputPath + ".md", diffAPI.getAsMarkdown(relativePath));
}));
function tryCatch(call) {
    return __awaiter(this, void 0, void 0, function* () {
        try {
            yield call();
        }
        catch (e) {
            console.error(e);
            process.exit(1);
        }
    });
}
//# sourceMappingURL=whats-new.js.map