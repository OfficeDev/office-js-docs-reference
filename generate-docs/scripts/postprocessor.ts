// declare var fileContents: string;

// const lines = fileContents
//     .replace(/\r\n/g, '\n')
//     .split("\n");

// const linesToKeep = [];
// for (let i = 0; i < lines.length; i++) {
//     let didChange = changeNamesIfRelevant(lines, i);
//     linesToKeep.push(lines[i]);
// }


// function changeNamesIfRelevant() {
//     //regex
//     // return true if it did in fact match and you ended up having to remove
// }

// // technique for updating structure of toc.yml
// const data: string = (jsyaml.safeLoad(fs.readFileSync(filename).toString()) as ISnippet).script.content;
// data.items.forEach((item, index) => {
//     item.name = item.name.substr(0, 1).toUpperCase() + item.name.substr(1);
//     if (true) {
//         item.items = item.items.items;
//     }
// }

// // step 1: delete everything except the 'overview' folder from the /docs folder
// export function deleteAllFilesExceptDotGit(path: string): void {
//     fs.readdirSync(path)
//         .filter(filename => filename !== ".git")
//         .forEach(filename => fs.removeSync(path + '/' + filename));
// }


// // step 2: copy all files/folders from /yaml folder to the /docs/docs-ref-autogen folder
// let options = { except: string[] }; //don't really need this...just specify string below in filter

// fs.readdirSync(tempCloneFolder)
//     .filter(filename => options.except.indexOf(filename) < 0)
//     .forEach(filename => {
//         fs.copySync(
//             tempCloneFolder + '/' + filename,
//             this.rootDestDir + '/' + filename,
//             COPY_OPTIONS
//         );
// });


// function fixTocFile() : void {
//     // read file
//     const file = fsx.readFileSync('tocORIG.yml');
//     let fileContent = file.toString();

//     // replace 
//     const result = fileContent
//         .replace(/(items:\r\n)(\s*\- name: Excel\r\n)(\s*items:\r\n)/, `$1`)
//         .replace(/(items:\r\n)(\s*\- name: OneNote\r\n)(\s*items:\r\n)/, `$1`)
//         .replace(/(items:\r\n)(\s*\- name: Visio\r\n)(\s*items:\r\n)/, `$1`)
//         .replace(/(items:\r\n)(\s*\- name: Word\r\n)(\s*items:\r\n)/, `$1`)
//         .replace(/(uid: outlook\r\n)(\s*items:\r\n)(\s*\- name: Office\r\n)(\s*items:\r\n)/, `$1$2`)
//         .replace(/name: excel/, `name: Excel`)
//         .replace(/name: office/, `name: Shared API`)
//         .replace(/name: onenote/, `name: OneNote`)
//         .replace(/name: outlook/, `name: Outlook`)
//         .replace(/name: visio/, `name: Visio`)
//         .replace(/name: word/, `name: Word`);
        
//     // TODO: at this point, the yaml contains extra indentation on some lines.
//     //       need to prettify the yaml (in 'result' string) before writing contents back to file.

//     // write file 
//     fsx.writeFileSync('toc.yml', result);
// }
