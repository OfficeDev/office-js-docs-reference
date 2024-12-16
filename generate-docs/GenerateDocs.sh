#!/bin/bash

while getopts b: flag
do
  case "${flag}" in
    b) bypassPrompt=${OPTARG};;
  esac
done

if [ -e "build-log.txt" ]; then
    rm build-log.txt
fi

if [ -e "build-errors.txt" ]; then
    rm build-errors.txt
fi

exec > >(tee -a build-log.txt) 2> >(tee -a build-errors.txt >&2)

if [ -d "node_modules" ]; then
    rm -rf "node_modules"
fi

if [ -d "scripts/node_modules" ]; then
    rm -rf "scripts/node_modules"
fi

if [ -d "tools/node_modules" ]; then
    rm -rf "tools/node_modules"
fi

if [ ! -d "json" ]; then
    mkdir json
fi

if [ ! -d "yaml" ]; then
    mkdir yaml
fi

npm install

pushd scripts
npm install
npm run build
node preprocessor.js $bypassPrompt
popd


pushd tools
mkdir tool-inputs
npm install
npm run build
node version-remover ../api-extractor-inputs-excel-release/Excel_online/excel.d.ts "ExcelApiOnline 1.1" ../api-extractor-inputs-excel-release/Excel_1_17/excel.d.ts
node version-remover ../api-extractor-inputs-excel-release/Excel_1_17/excel.d.ts "ExcelApi 1.17" ../api-extractor-inputs-excel-release/Excel_1_16/excel.d.ts
node version-remover ../api-extractor-inputs-excel-release/Excel_1_16/excel.d.ts "ExcelApi 1.16" ../api-extractor-inputs-excel-release/Excel_1_15/excel.d.ts
node version-remover ../api-extractor-inputs-excel-release/Excel_1_15/excel.d.ts "ExcelApi 1.15" ../api-extractor-inputs-excel-release/Excel_1_14/excel.d.ts
node version-remover ../api-extractor-inputs-excel-release/Excel_1_14/excel.d.ts "ExcelApi 1.14" ../api-extractor-inputs-excel-release/Excel_1_13/excel.d.ts
node version-remover ../api-extractor-inputs-excel-release/Excel_1_13/excel.d.ts "ExcelApi 1.13" ../api-extractor-inputs-excel-release/Excel_1_12/excel.d.ts
node version-remover ../api-extractor-inputs-excel-release/Excel_1_12/excel.d.ts "ExcelApi 1.12" ../api-extractor-inputs-excel-release/Excel_1_11/excel.d.ts
node version-remover ../api-extractor-inputs-excel-release/Excel_1_11/excel.d.ts "ExcelApi 1.11" ../api-extractor-inputs-excel-release/Excel_1_10/excel.d.ts
node version-remover ../api-extractor-inputs-excel-release/Excel_1_10/excel.d.ts "ExcelApi 1.10" ../api-extractor-inputs-excel-release/Excel_1_9/excel.d.ts
node version-remover ../api-extractor-inputs-excel-release/Excel_1_9/excel.d.ts "ExcelApi 1.9" ../api-extractor-inputs-excel-release/Excel_1_8/excel.d.ts
node version-remover ../api-extractor-inputs-excel-release/Excel_1_8/excel.d.ts "ExcelApi 1.8" ../api-extractor-inputs-excel-release/Excel_1_7/excel.d.ts
node version-remover ../api-extractor-inputs-excel-release/Excel_1_7/excel.d.ts "ExcelApi 1.7" ../api-extractor-inputs-excel-release/Excel_1_6/excel.d.ts
node version-remover ../api-extractor-inputs-excel-release/Excel_1_6/excel.d.ts "ExcelApi 1.6" ../api-extractor-inputs-excel-release/Excel_1_5/excel.d.ts
node version-remover ../api-extractor-inputs-excel-release/Excel_1_5/excel.d.ts "ExcelApi 1.5" ../api-extractor-inputs-excel-release/Excel_1_4/excel.d.ts
node version-remover ../api-extractor-inputs-excel-release/Excel_1_4/excel.d.ts "ExcelApi 1.4" ../api-extractor-inputs-excel-release/Excel_1_3/excel.d.ts
node version-remover ../api-extractor-inputs-excel-release/Excel_1_3/excel.d.ts "ExcelApi 1.3" ../api-extractor-inputs-excel-release/Excel_1_2/excel.d.ts
node version-remover ../api-extractor-inputs-excel-release/Excel_1_2/excel.d.ts "ExcelApi 1.2" ../api-extractor-inputs-excel-release/Excel_1_1/excel.d.ts
node version-remover ../api-extractor-inputs-excel-release/Excel_1_1/excel.d.ts "ExcelApi 1.1" ./tool-inputs/excel-base.d.ts

node version-remover ../api-extractor-inputs-outlook-release/outlook_1_14/outlook.d.ts "Mailbox 1.14" ../api-extractor-inputs-outlook-release/outlook_1_13/outlook.d.ts
node ../scripts/versioned-dts-cleanup ../api-extractor-inputs-outlook-release/outlook_1_13/outlook.d.ts Outlook 1.13
node version-remover ../api-extractor-inputs-outlook-release/outlook_1_13/outlook.d.ts "Mailbox 1.13" ../api-extractor-inputs-outlook-release/outlook_1_12/outlook.d.ts
node ../scripts/versioned-dts-cleanup ../api-extractor-inputs-outlook-release/outlook_1_12/outlook.d.ts Outlook 1.12
node version-remover ../api-extractor-inputs-outlook-release/outlook_1_12/outlook.d.ts "Mailbox 1.12" ../api-extractor-inputs-outlook-release/outlook_1_11/outlook.d.ts
node ../scripts/versioned-dts-cleanup ../api-extractor-inputs-outlook-release/outlook_1_11/outlook.d.ts Outlook 1.11
node version-remover ../api-extractor-inputs-outlook-release/outlook_1_11/outlook.d.ts "Mailbox 1.11" ../api-extractor-inputs-outlook-release/outlook_1_10/outlook.d.ts
node ../scripts/versioned-dts-cleanup ../api-extractor-inputs-outlook-release/outlook_1_10/outlook.d.ts Outlook 1.10
node version-remover ../api-extractor-inputs-outlook-release/outlook_1_10/outlook.d.ts "Mailbox 1.10" ../api-extractor-inputs-outlook-release/outlook_1_9/outlook.d.ts
node ../scripts/versioned-dts-cleanup ../api-extractor-inputs-outlook-release/outlook_1_9/outlook.d.ts Outlook 1.9
node version-remover ../api-extractor-inputs-outlook-release/outlook_1_9/outlook.d.ts "Mailbox 1.9" ../api-extractor-inputs-outlook-release/outlook_1_8/outlook.d.ts
node ../scripts/versioned-dts-cleanup ../api-extractor-inputs-outlook-release/outlook_1_8/outlook.d.ts Outlook 1.8
node version-remover ../api-extractor-inputs-outlook-release/outlook_1_8/outlook.d.ts "Mailbox 1.8" ../api-extractor-inputs-outlook-release/outlook_1_7/outlook.d.ts
node ../scripts/versioned-dts-cleanup ../api-extractor-inputs-outlook-release/outlook_1_7/outlook.d.ts Outlook 1.7
node version-remover ../api-extractor-inputs-outlook-release/outlook_1_7/outlook.d.ts "Mailbox 1.7" ../api-extractor-inputs-outlook-release/outlook_1_6/outlook.d.ts
node ../scripts/versioned-dts-cleanup ../api-extractor-inputs-outlook-release/outlook_1_6/outlook.d.ts Outlook 1.6
node version-remover ../api-extractor-inputs-outlook-release/outlook_1_6/outlook.d.ts "Mailbox 1.6" ../api-extractor-inputs-outlook-release/outlook_1_5/outlook.d.ts
node ../scripts/versioned-dts-cleanup ../api-extractor-inputs-outlook-release/outlook_1_5/outlook.d.ts Outlook 1.5
node version-remover ../api-extractor-inputs-outlook-release/outlook_1_5/outlook.d.ts "Mailbox 1.5" ../api-extractor-inputs-outlook-release/outlook_1_4/outlook.d.ts
node ../scripts/versioned-dts-cleanup ../api-extractor-inputs-outlook-release/outlook_1_4/outlook.d.ts Outlook 1.4
node version-remover ../api-extractor-inputs-outlook-release/outlook_1_4/outlook.d.ts "Mailbox 1.4" ../api-extractor-inputs-outlook-release/outlook_1_3/outlook.d.ts
node ../scripts/versioned-dts-cleanup ../api-extractor-inputs-outlook-release/outlook_1_3/outlook.d.ts Outlook 1.3
node version-remover ../api-extractor-inputs-outlook-release/outlook_1_3/outlook.d.ts "Mailbox 1.3" ../api-extractor-inputs-outlook-release/outlook_1_2/outlook.d.ts
node ../scripts/versioned-dts-cleanup ../api-extractor-inputs-outlook-release/outlook_1_2/outlook.d.ts Outlook 1.2
node version-remover ../api-extractor-inputs-outlook-release/outlook_1_2/outlook.d.ts "Mailbox 1.2" ../api-extractor-inputs-outlook-release/outlook_1_1/outlook.d.ts
node ../scripts/versioned-dts-cleanup ../api-extractor-inputs-outlook-release/outlook_1_1/outlook.d.ts Outlook 1.1
node version-remover ../api-extractor-inputs-outlook-release/outlook_1_1/outlook.d.ts "Mailbox 1.1" ./tool-inputs/outlook-base.d.ts

node version-remover ../api-extractor-inputs-powerpoint-release/powerpoint_1_7/powerpoint.d.ts "PowerPointApi 1.7" ../api-extractor-inputs-powerpoint-release/powerpoint_1_6/powerpoint.d.ts
node version-remover ../api-extractor-inputs-powerpoint-release/powerpoint_1_6/powerpoint.d.ts "PowerPointApi 1.6" ../api-extractor-inputs-powerpoint-release/powerpoint_1_5/powerpoint.d.ts
node version-remover ../api-extractor-inputs-powerpoint-release/powerpoint_1_5/powerpoint.d.ts "PowerPointApi 1.5" ../api-extractor-inputs-powerpoint-release/powerpoint_1_4/powerpoint.d.ts
node version-remover ../api-extractor-inputs-powerpoint-release/powerpoint_1_4/powerpoint.d.ts "PowerPointApi 1.4" ../api-extractor-inputs-powerpoint-release/powerpoint_1_3/powerpoint.d.ts
node version-remover ../api-extractor-inputs-powerpoint-release/powerpoint_1_3/powerpoint.d.ts "PowerPointApi 1.3" ../api-extractor-inputs-powerpoint-release/powerpoint_1_2/powerpoint.d.ts
node version-remover ../api-extractor-inputs-powerpoint-release/powerpoint_1_2/powerpoint.d.ts "PowerPointApi 1.2" ../api-extractor-inputs-powerpoint-release/powerpoint_1_1/powerpoint.d.ts
node version-remover ../api-extractor-inputs-powerpoint-release/powerpoint_1_1/powerpoint.d.ts "PowerPointApi 1.1" ./tool-inputs/powerpoint-base.d.ts

node version-remover ../api-extractor-inputs-word-release/word_online/word-init.d.ts "WordApiOnline 1.1" ../api-extractor-inputs-word-release/word_desktop_1_1/word-desktop1.d.ts
node version-remover ../api-extractor-inputs-word-release/word_desktop_1_1/word-desktop1.d.ts "WordApi 1.9" ../api-extractor-inputs-word-release/word_desktop_1_1/word-desktop2.d.ts
node version-remover ../api-extractor-inputs-word-release/word_desktop_1_1/word-desktop2.d.ts "WordApiHiddenDocument 1.5" ../api-extractor-inputs-word-release/word_desktop_1_1/word-desktop3.d.ts
node version-remover ../api-extractor-inputs-word-release/word_desktop_1_1/word-desktop3.d.ts "WordApiHiddenDocument 1.4" ../api-extractor-inputs-word-release/word_desktop_1_1/word-desktop4.d.ts
node version-remover ../api-extractor-inputs-word-release/word_desktop_1_1/word-desktop4.d.ts "WordApiHiddenDocument 1.3" ../api-extractor-inputs-word-release/word_desktop_1_1/word.d.ts
node version-remover ../api-extractor-inputs-word-release/word_online/word-init.d.ts "WordApiDesktop 1.1" ../api-extractor-inputs-word-release/word_online/word-online1.d.ts
node version-remover ../api-extractor-inputs-word-release/word_online/word-online1.d.ts "WordApiHiddenDocument 1.5" ../api-extractor-inputs-word-release/word_online/word-online2.d.ts
node version-remover ../api-extractor-inputs-word-release/word_online/word-online2.d.ts "WordApiHiddenDocument 1.4" ../api-extractor-inputs-word-release/word_online/word-online3.d.ts
node version-remover ../api-extractor-inputs-word-release/word_online/word-online3.d.ts "WordApiHiddenDocument 1.3" ../api-extractor-inputs-word-release/word_online/word.d.ts
node version-remover ../api-extractor-inputs-word-release/word_desktop_1_1/word-desktop2.d.ts "WordApiDesktop 1.1" ../api-extractor-inputs-word-release/word_1_5_hidden_document/word-desktop1.d.ts
node version-remover ../api-extractor-inputs-word-release/word_1_5_hidden_document/word-desktop1.d.ts "WordApi 1.8" ../api-extractor-inputs-word-release/word_1_5_hidden_document/word-desktop2.d.ts
node version-remover ../api-extractor-inputs-word-release/word_1_5_hidden_document/word-desktop2.d.ts "WordApi 1.7" ../api-extractor-inputs-word-release/word_1_5_hidden_document/word-desktop3.d.ts
node version-remover ../api-extractor-inputs-word-release/word_1_5_hidden_document/word-desktop3.d.ts "WordApi 1.6" ../api-extractor-inputs-word-release/word_1_5_hidden_document/word.d.ts
node version-remover ../api-extractor-inputs-word-release/word_1_5_hidden_document/word.d.ts "WordApiHiddenDocument 1.5" ../api-extractor-inputs-word-release/word_1_4_hidden_document/word-desktop1.d.ts
node version-remover ../api-extractor-inputs-word-release/word_1_4_hidden_document/word-desktop1.d.ts "WordApi 1.5" ../api-extractor-inputs-word-release/word_1_4_hidden_document/word.d.ts
node version-remover ../api-extractor-inputs-word-release/word_1_4_hidden_document/word.d.ts "WordApiHiddenDocument 1.4" ../api-extractor-inputs-word-release/word_1_3_hidden_document/word-desktop1.d.ts
node version-remover ../api-extractor-inputs-word-release/word_1_3_hidden_document/word-desktop1.d.ts "WordApi 1.4" ../api-extractor-inputs-word-release/word_1_3_hidden_document/word.d.ts
node version-remover ../api-extractor-inputs-word-release/word_online/word.d.ts "WordApiOnline 1.1" ../api-extractor-inputs-word-release/word_1_9/word.d.ts
node version-remover ../api-extractor-inputs-word-release/word_1_9/word.d.ts "WordApi 1.9" ../api-extractor-inputs-word-release/word_1_8/word.d.ts
node version-remover ../api-extractor-inputs-word-release/word_1_8/word.d.ts "WordApi 1.8" ../api-extractor-inputs-word-release/word_1_7/word.d.ts
node version-remover ../api-extractor-inputs-word-release/word_1_7/word.d.ts "WordApi 1.7" ../api-extractor-inputs-word-release/word_1_6/word.d.ts
node version-remover ../api-extractor-inputs-word-release/word_1_6/word.d.ts "WordApi 1.6" ../api-extractor-inputs-word-release/word_1_5/word.d.ts
node version-remover ../api-extractor-inputs-word-release/word_1_5/word.d.ts "WordApi 1.5" ../api-extractor-inputs-word-release/word_1_4/word.d.ts
node version-remover ../api-extractor-inputs-word-release/word_1_4/word.d.ts "WordApi 1.4" ../api-extractor-inputs-word-release/word_1_3/word.d.ts
node version-remover ../api-extractor-inputs-word-release/word_1_3/word.d.ts "WordApi 1.3" ../api-extractor-inputs-word-release/word_1_2/word.d.ts
node version-remover ../api-extractor-inputs-word-release/word_1_2/word.d.ts "WordApi 1.2" ../api-extractor-inputs-word-release/word_1_1/word.d.ts
node version-remover ../api-extractor-inputs-word-release/word_1_1/word.d.ts "WordApi 1.1" ./tool-inputs/word-base.d.ts


node whats-new excel ../api-extractor-inputs-excel/excel.d.ts ../api-extractor-inputs-excel-release/Excel_online/excel.d.ts ../../docs/includes/excel-preview
node whats-new excel ../api-extractor-inputs-excel-release/Excel_online/excel.d.ts ../api-extractor-inputs-excel-release/Excel_1_17/excel.d.ts ../../docs/includes/excel-online
node whats-new excel ../api-extractor-inputs-excel-release/Excel_1_17/excel.d.ts ../api-extractor-inputs-excel-release/Excel_1_16/excel.d.ts ../../docs/includes/excel-1_17
node whats-new excel ../api-extractor-inputs-excel-release/Excel_1_16/excel.d.ts ../api-extractor-inputs-excel-release/Excel_1_15/excel.d.ts ../../docs/includes/excel-1_16
node whats-new excel ../api-extractor-inputs-excel-release/Excel_1_15/excel.d.ts ../api-extractor-inputs-excel-release/Excel_1_14/excel.d.ts ../../docs/includes/excel-1_15
node whats-new excel ../api-extractor-inputs-excel-release/Excel_1_14/excel.d.ts ../api-extractor-inputs-excel-release/Excel_1_13/excel.d.ts ../../docs/includes/excel-1_14
node whats-new excel ../api-extractor-inputs-excel-release/Excel_1_13/excel.d.ts ../api-extractor-inputs-excel-release/Excel_1_12/excel.d.ts ../../docs/includes/excel-1_13
node whats-new excel ../api-extractor-inputs-excel-release/Excel_1_12/excel.d.ts ../api-extractor-inputs-excel-release/Excel_1_11/excel.d.ts ../../docs/includes/excel-1_12
node whats-new excel ../api-extractor-inputs-excel-release/Excel_1_11/excel.d.ts ../api-extractor-inputs-excel-release/Excel_1_10/excel.d.ts ../../docs/includes/excel-1_11
node whats-new excel ../api-extractor-inputs-excel-release/Excel_1_10/excel.d.ts ../api-extractor-inputs-excel-release/Excel_1_9/excel.d.ts ../../docs/includes/excel-1_10
node whats-new excel ../api-extractor-inputs-excel-release/Excel_1_9/excel.d.ts ../api-extractor-inputs-excel-release/Excel_1_8/excel.d.ts ../../docs/includes/excel-1_9
node whats-new excel ../api-extractor-inputs-excel-release/Excel_1_8/excel.d.ts ../api-extractor-inputs-excel-release/Excel_1_7/excel.d.ts ../../docs/includes/excel-1_8
node whats-new excel ../api-extractor-inputs-excel-release/Excel_1_7/excel.d.ts ../api-extractor-inputs-excel-release/Excel_1_6/excel.d.ts ../../docs/includes/excel-1_7
node whats-new excel ../api-extractor-inputs-excel-release/Excel_1_6/excel.d.ts ../api-extractor-inputs-excel-release/Excel_1_5/excel.d.ts ../../docs/includes/excel-1_6
node whats-new excel ../api-extractor-inputs-excel-release/Excel_1_5/excel.d.ts ../api-extractor-inputs-excel-release/Excel_1_4/excel.d.ts ../../docs/includes/excel-1_5
node whats-new excel ../api-extractor-inputs-excel-release/Excel_1_4/excel.d.ts ../api-extractor-inputs-excel-release/Excel_1_3/excel.d.ts ../../docs/includes/excel-1_4
node whats-new excel ../api-extractor-inputs-excel-release/Excel_1_3/excel.d.ts ../api-extractor-inputs-excel-release/Excel_1_2/excel.d.ts ../../docs/includes/excel-1_3
node whats-new excel ../api-extractor-inputs-excel-release/Excel_1_2/excel.d.ts ../api-extractor-inputs-excel-release/Excel_1_1/excel.d.ts ../../docs/includes/excel-1_2
node whats-new excel ../api-extractor-inputs-excel-release/Excel_1_1/excel.d.ts ./tool-inputs/excel-base.d.ts ../../docs/includes/excel-1_1

node whats-new outlook ../api-extractor-inputs-outlook/outlook.d.ts ../api-extractor-inputs-outlook-release/outlook_1_14/outlook.d.ts ../../docs/includes/outlook-preview
node whats-new outlook ../api-extractor-inputs-outlook-release/outlook_1_14/outlook.d.ts ../api-extractor-inputs-outlook-release/outlook_1_13/outlook.d.ts ../../docs/includes/outlook-1_14
node whats-new outlook ../api-extractor-inputs-outlook-release/outlook_1_13/outlook.d.ts ../api-extractor-inputs-outlook-release/outlook_1_12/outlook.d.ts ../../docs/includes/outlook-1_13
node whats-new outlook ../api-extractor-inputs-outlook-release/outlook_1_12/outlook.d.ts ../api-extractor-inputs-outlook-release/outlook_1_11/outlook.d.ts ../../docs/includes/outlook-1_12
node whats-new outlook ../api-extractor-inputs-outlook-release/outlook_1_11/outlook.d.ts ../api-extractor-inputs-outlook-release/outlook_1_10/outlook.d.ts ../../docs/includes/outlook-1_11
node whats-new outlook ../api-extractor-inputs-outlook-release/outlook_1_10/outlook.d.ts ../api-extractor-inputs-outlook-release/outlook_1_9/outlook.d.ts ../../docs/includes/outlook-1_10
node whats-new outlook ../api-extractor-inputs-outlook-release/outlook_1_9/outlook.d.ts ../api-extractor-inputs-outlook-release/outlook_1_8/outlook.d.ts ../../docs/includes/outlook-1_9
node whats-new outlook ../api-extractor-inputs-outlook-release/outlook_1_8/outlook.d.ts ../api-extractor-inputs-outlook-release/outlook_1_7/outlook.d.ts ../../docs/includes/outlook-1_8
node whats-new outlook ../api-extractor-inputs-outlook-release/outlook_1_7/outlook.d.ts ../api-extractor-inputs-outlook-release/outlook_1_6/outlook.d.ts ../../docs/includes/outlook-1_7
node whats-new outlook ../api-extractor-inputs-outlook-release/outlook_1_6/outlook.d.ts ../api-extractor-inputs-outlook-release/outlook_1_5/outlook.d.ts ../../docs/includes/outlook-1_6
node whats-new outlook ../api-extractor-inputs-outlook-release/outlook_1_5/outlook.d.ts ../api-extractor-inputs-outlook-release/outlook_1_4/outlook.d.ts ../../docs/includes/outlook-1_5
node whats-new outlook ../api-extractor-inputs-outlook-release/outlook_1_4/outlook.d.ts ../api-extractor-inputs-outlook-release/outlook_1_3/outlook.d.ts ../../docs/includes/outlook-1_4
node whats-new outlook ../api-extractor-inputs-outlook-release/outlook_1_3/outlook.d.ts ../api-extractor-inputs-outlook-release/outlook_1_2/outlook.d.ts ../../docs/includes/outlook-1_3
node whats-new outlook ../api-extractor-inputs-outlook-release/outlook_1_2/outlook.d.ts ../api-extractor-inputs-outlook-release/outlook_1_1/outlook.d.ts ../../docs/includes/outlook-1_2
node whats-new outlook ../api-extractor-inputs-outlook-release/outlook_1_1/outlook.d.ts ./tool-inputs/outlook-base.d.ts ../../docs/includes/outlook-1_1

node whats-new powerpoint ../api-extractor-inputs-powerpoint/powerpoint.d.ts ../api-extractor-inputs-powerpoint-release/powerpoint_1_7/powerpoint.d.ts ../../docs/includes/powerpoint-preview
node whats-new powerpoint ../api-extractor-inputs-powerpoint-release/powerpoint_1_7/powerpoint.d.ts ../api-extractor-inputs-powerpoint-release/powerpoint_1_6/powerpoint.d.ts ../../docs/includes/powerpoint-1_7
node whats-new powerpoint ../api-extractor-inputs-powerpoint-release/powerpoint_1_6/powerpoint.d.ts ../api-extractor-inputs-powerpoint-release/powerpoint_1_5/powerpoint.d.ts ../../docs/includes/powerpoint-1_6
node whats-new powerpoint ../api-extractor-inputs-powerpoint-release/powerpoint_1_5/powerpoint.d.ts ../api-extractor-inputs-powerpoint-release/powerpoint_1_4/powerpoint.d.ts ../../docs/includes/powerpoint-1_5
node whats-new powerpoint ../api-extractor-inputs-powerpoint-release/powerpoint_1_4/powerpoint.d.ts ../api-extractor-inputs-powerpoint-release/powerpoint_1_3/powerpoint.d.ts ../../docs/includes/powerpoint-1_4
node whats-new powerpoint ../api-extractor-inputs-powerpoint-release/powerpoint_1_3/powerpoint.d.ts ../api-extractor-inputs-powerpoint-release/powerpoint_1_2/powerpoint.d.ts ../../docs/includes/powerpoint-1_3
node whats-new powerpoint ../api-extractor-inputs-powerpoint-release/powerpoint_1_2/powerpoint.d.ts ../api-extractor-inputs-powerpoint-release/powerpoint_1_1/powerpoint.d.ts ../../docs/includes/powerpoint-1_2
node whats-new powerpoint ../api-extractor-inputs-powerpoint-release/powerpoint_1_1/powerpoint.d.ts ./tool-inputs/powerpoint-base.d.ts ../../docs/includes/powerpoint-1_1

node whats-new word ../api-extractor-inputs-word/word.d.ts ../api-extractor-inputs-word-release/word_online/word-init.d.ts ../../docs/includes/word-preview
node whats-new word ../api-extractor-inputs-word-release/word_online/word.d.ts ../api-extractor-inputs-word-release/word_1_9/word.d.ts ../../docs/includes/word-online
node whats-new word ../api-extractor-inputs-word-release/word_desktop_1_1/word.d.ts ../api-extractor-inputs-word-release/word_1_8/word.d.ts ../../docs/includes/word-desktop-1.1
node whats-new word ../api-extractor-inputs-word-release/word_1_5_hidden_document/word.d.ts ../api-extractor-inputs-word-release/word_1_4_hidden_document/word-desktop1.d.ts ../../docs/includes/word-1_5_hidden_document
node whats-new word ../api-extractor-inputs-word-release/word_1_4_hidden_document/word.d.ts ../api-extractor-inputs-word-release/word_1_3_hidden_document/word-desktop1.d.ts ../../docs/includes/word-1_4_hidden_document
node whats-new word ../api-extractor-inputs-word-release/word_1_3_hidden_document/word.d.ts ../api-extractor-inputs-word-release/word_1_3/word.d.ts ../../docs/includes/word-1_3_hidden_document
node whats-new word ../api-extractor-inputs-word-release/word_1_9/word.d.ts ../api-extractor-inputs-word-release/word_1_8/word.d.ts ../../docs/includes/word-1_9
node whats-new word ../api-extractor-inputs-word-release/word_1_8/word.d.ts ../api-extractor-inputs-word-release/word_1_7/word.d.ts ../../docs/includes/word-1_8
node whats-new word ../api-extractor-inputs-word-release/word_1_7/word.d.ts ../api-extractor-inputs-word-release/word_1_6/word.d.ts ../../docs/includes/word-1_7
node whats-new word ../api-extractor-inputs-word-release/word_1_6/word.d.ts ../api-extractor-inputs-word-release/word_1_5/word.d.ts ../../docs/includes/word-1_6
node whats-new word ../api-extractor-inputs-word-release/word_1_5/word.d.ts ../api-extractor-inputs-word-release/word_1_4/word.d.ts ../../docs/includes/word-1_5
node whats-new word ../api-extractor-inputs-word-release/word_1_4/word.d.ts ../api-extractor-inputs-word-release/word_1_3/word.d.ts ../../docs/includes/word-1_4
node whats-new word ../api-extractor-inputs-word-release/word_1_3/word.d.ts ../api-extractor-inputs-word-release/word_1_2/word.d.ts ../../docs/includes/word-1_3
node whats-new word ../api-extractor-inputs-word-release/word_1_2/word.d.ts ../api-extractor-inputs-word-release/word_1_1/word.d.ts ../../docs/includes/word-1_2
node whats-new word ../api-extractor-inputs-word-release/word_1_1/word.d.ts ./tool-inputs/word-base.d.ts ../../docs/includes/word-1_1

popd

if [ ! -d "json/office" ]; then
    echo Running API Extractor for Office preview.
    pushd api-extractor-inputs-office
    ../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/office-release" ]; then
    echo Running API Extractor for Office release.
    pushd api-extractor-inputs-office-release
    ../node_modules/.bin/api-extractor run
    popd
fi

if [ ! -d "json/excel" ]; then
    echo Running API Extractor for Excel preview.
    pushd api-extractor-inputs-excel
    ../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/excel_online" ]; then
    echo Running API Extractor for Excel online.
    pushd api-extractor-inputs-excel-release/excel_online
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/excel_1_17" ]; then
    echo Running API Extractor for Excel 1.17.
    pushd api-extractor-inputs-excel-release/excel_1_17
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/excel_1_16" ]; then
    echo Running API Extractor for Excel 1.16.
    pushd api-extractor-inputs-excel-release/excel_1_16
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/excel_1_15" ]; then
    echo Running API Extractor for Excel 1.15.
    pushd api-extractor-inputs-excel-release/excel_1_15
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/excel_1_14" ]; then
    echo Running API Extractor for Excel 1.14.
    pushd api-extractor-inputs-excel-release/excel_1_14
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/excel_1_13" ]; then
    echo Running API Extractor for Excel 1.13.
    pushd api-extractor-inputs-excel-release/excel_1_13
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/excel_1_12" ]; then
    echo Running API Extractor for Excel 1.12.
    pushd api-extractor-inputs-excel-release/excel_1_12
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/excel_1_11" ]; then
    echo Running API Extractor for Excel 1.11.
    pushd api-extractor-inputs-excel-release/excel_1_11
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/excel_1_10" ]; then
    echo Running API Extractor for Excel 1.10.
    pushd api-extractor-inputs-excel-release/excel_1_10
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/excel_1_9" ]; then
    echo Running API Extractor for Excel 1.9.
    pushd api-extractor-inputs-excel-release/excel_1_9
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/excel_1_8" ]; then
    echo Running API Extractor for Excel 1.8.
    pushd api-extractor-inputs-excel-release/excel_1_8
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/excel_1_7" ]; then
    echo Running API Extractor for Excel 1.7.
    pushd api-extractor-inputs-excel-release/excel_1_7
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/excel_1_6" ]; then
    echo Running API Extractor for Excel 1.6.
    pushd api-extractor-inputs-excel-release/excel_1_6
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/excel_1_5" ]; then
    echo Running API Extractor for Excel 1.5.
    pushd api-extractor-inputs-excel-release/excel_1_5
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/excel_1_4" ]; then
    echo Running API Extractor for Excel 1.4.
    pushd api-extractor-inputs-excel-release/excel_1_4
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/excel_1_3" ]; then
    echo Running API Extractor for Excel 1.3.
    pushd api-extractor-inputs-excel-release/excel_1_3
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/excel_1_2" ]; then
    echo Running API Extractor for Excel 1.2.
    pushd api-extractor-inputs-excel-release/excel_1_2
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/excel_1_1" ]; then
    echo Running API Extractor for Excel 1.1.
    pushd api-extractor-inputs-excel-release/excel_1_1
    ../../node_modules/.bin/api-extractor run
    popd
fi

if [ ! -d "json/onenote" ]; then
    echo Running API Extractor for OneNote.
    pushd api-extractor-inputs-onenote
    ../node_modules/.bin/api-extractor run
    popd
fi

if [ ! -d "json/outlook" ]; then
    echo Running API Extractor for Outlook preview.
    pushd api-extractor-inputs-outlook
    ../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/outlook_1_14" ]; then
    echo Running API Extractor for Outlook 1.14.
    pushd api-extractor-inputs-outlook-release/outlook_1_14
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/outlook_1_13" ]; then
    echo Running API Extractor for Outlook 1.13.
    pushd api-extractor-inputs-outlook-release/outlook_1_13
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/outlook_1_12" ]; then
    echo Running API Extractor for Outlook 1.12.
    pushd api-extractor-inputs-outlook-release/outlook_1_12
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/outlook_1_11" ]; then
    echo Running API Extractor for Outlook 1.11.
    pushd api-extractor-inputs-outlook-release/outlook_1_11
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/outlook_1_10" ]; then
    echo Running API Extractor for Outlook 1.10.
    pushd api-extractor-inputs-outlook-release/outlook_1_10
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/outlook_1_9" ]; then
    echo Running API Extractor for Outlook 1.9.
    pushd api-extractor-inputs-outlook-release/outlook_1_9
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/outlook_1_8" ]; then
    echo Running API Extractor for Outlook 1.8.
    pushd api-extractor-inputs-outlook-release/outlook_1_8
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/outlook_1_7" ]; then
    echo Running API Extractor for Outlook 1.7.
    pushd api-extractor-inputs-outlook-release/outlook_1_7
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/outlook_1_6" ]; then
    echo Running API Extractor for Outlook 1.6.
    pushd api-extractor-inputs-outlook-release/outlook_1_6
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/outlook_1_5" ]; then
    echo Running API Extractor for Outlook 1.5.
    pushd api-extractor-inputs-outlook-release/outlook_1_5
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/outlook_1_4" ]; then
    echo Running API Extractor for Outlook 1.4.
    pushd api-extractor-inputs-outlook-release/outlook_1_4
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/outlook_1_3" ]; then
    echo Running API Extractor for Outlook 1.3.
    pushd api-extractor-inputs-outlook-release/outlook_1_3
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/outlook_1_2" ]; then
    echo Running API Extractor for Outlook 1.2.
    pushd api-extractor-inputs-outlook-release/outlook_1_2
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/outlook_1_1" ]; then
    echo Running API Extractor for Outlook 1.1.
    pushd api-extractor-inputs-outlook-release/outlook_1_1
    ../../node_modules/.bin/api-extractor run
    popd
fi

if [ ! -d "json/powerpoint" ]; then
    echo Running API Extractor for PowerPoint preview.
    pushd api-extractor-inputs-powerpoint
    ../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/powerpoint_1_7" ]; then
    echo Running API Extractor for PowerPoint 1.7.
    pushd api-extractor-inputs-powerpoint-release/PowerPoint_1_7
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/powerpoint_1_6" ]; then
    echo Running API Extractor for PowerPoint 1.6.
    pushd api-extractor-inputs-powerpoint-release/PowerPoint_1_6
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/powerpoint_1_5" ]; then
    echo Running API Extractor for PowerPoint 1.5.
    pushd api-extractor-inputs-powerpoint-release/PowerPoint_1_5
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/powerpoint_1_4" ]; then
    echo Running API Extractor for PowerPoint 1.4.
    pushd api-extractor-inputs-powerpoint-release/PowerPoint_1_4
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/powerpoint_1_3" ]; then
    echo Running API Extractor for PowerPoint 1.3.
    pushd api-extractor-inputs-powerpoint-release/PowerPoint_1_3
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/powerpoint_1_2" ]; then
    echo Running API Extractor for PowerPoint 1.2.
    pushd api-extractor-inputs-powerpoint-release/PowerPoint_1_2
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/powerpoint_1_1" ]; then
    echo Running API Extractor for PowerPoint 1.1.
    pushd api-extractor-inputs-powerpoint-release/PowerPoint_1_1
    ../../node_modules/.bin/api-extractor run
    popd
fi

if [ ! -d "json/visio" ]; then
    echo Running API Extractor for Visio.
    pushd api-extractor-inputs-visio
    ../node_modules/.bin/api-extractor run
    popd
fi

if [ ! -d "json/word" ]; then
    echo Running API Extractor for Word preview.
    pushd api-extractor-inputs-word
    ../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/word_online" ]; then
    echo Running API Extractor for Word online.
    pushd api-extractor-inputs-word-release/word_online
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/word_desktop_1_1" ]; then
    echo Running API Extractor for Word desktop 1.1.
    pushd api-extractor-inputs-word-release/word_desktop_1_1
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/word_1_5_hidden_document" ]; then
    echo Running API Extractor for Word desktop hidden document 1.5.
    pushd api-extractor-inputs-word-release/word_1_5_hidden_document
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/word_1_4_hidden_document" ]; then
    echo Running API Extractor for Word desktop hidden document 1.4.
    pushd api-extractor-inputs-word-release/word_1_4_hidden_document
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/word_1_3_hidden_document" ]; then
    echo Running API Extractor for Word desktop hidden document 1.3.
    pushd api-extractor-inputs-word-release/word_1_3_hidden_document
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/word_1_9" ]; then
    echo Running API Extractor for Word 1.9.
    pushd api-extractor-inputs-word-release/word_1_9
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/word_1_8" ]; then
    echo Running API Extractor for Word 1.8.
    pushd api-extractor-inputs-word-release/word_1_8
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/word_1_7" ]; then
    echo Running API Extractor for Word 1.7.
    pushd api-extractor-inputs-word-release/word_1_7
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/word_1_6" ]; then
    echo Running API Extractor for Word 1.6.
    pushd api-extractor-inputs-word-release/word_1_6
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/word_1_5" ]; then
    echo Running API Extractor for Word 1.5.
    pushd api-extractor-inputs-word-release/word_1_5
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/word_1_4" ]; then
    echo Running API Extractor for Word 1.4.
    pushd api-extractor-inputs-word-release/word_1_4
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/word_1_3" ]; then
    echo Running API Extractor for Word 1.3.
    pushd api-extractor-inputs-word-release/word_1_3
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/word_1_2" ]; then
    echo Running API Extractor for Word 1.2.
    pushd api-extractor-inputs-word-release/word_1_2
    ../../node_modules/.bin/api-extractor run
    popd
fi
if [ ! -d "json/word_1_1" ]; then
    echo Running API Extractor for Word 1.1.
    pushd api-extractor-inputs-word-release/word_1_1
    ../../node_modules/.bin/api-extractor run
    popd
fi

echo Running API Extractor for Custom Functions.
pushd api-extractor-inputs-custom-functions-runtime
../node_modules/.bin/api-extractor run
popd

echo Running API Extractor for Office Runtime.
pushd api-extractor-inputs-office-runtime
../node_modules/.bin/api-extractor run
popd

pushd scripts
node midprocessor.js
popd


if [ ! -d "yaml/office" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/office --output-folder ./yaml/office --office
fi
if [ ! -d "yaml/office_release" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/office_release --output-folder ./yaml/office_release --office 2>/dev/null
fi

if [ ! -d "yaml/office-runtime" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/office-runtime --output-folder ./yaml/office-runtime --office
fi

if [ ! -d "yaml/excel" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/excel --output-folder ./yaml/excel --office
fi
if [ ! -d "yaml/excel_1_1" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/excel_1_1 --output-folder ./yaml/excel_1_1 --office 2>/dev/null
fi
if [ ! -d "yaml/excel_1_2" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/excel_1_2 --output-folder ./yaml/excel_1_2 --office 2>/dev/null
fi
if [ ! -d "yaml/excel_1_3" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/excel_1_3 --output-folder ./yaml/excel_1_3 --office 2>/dev/null
fi
if [ ! -d "yaml/excel_1_4" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/excel_1_4 --output-folder ./yaml/excel_1_4 --office 2>/dev/null
fi
if [ ! -d "yaml/excel_1_5" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/excel_1_5 --output-folder ./yaml/excel_1_5 --office 2>/dev/null
fi
if [ ! -d "yaml/excel_1_6" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/excel_1_6 --output-folder ./yaml/excel_1_6 --office 2>/dev/null
fi
if [ ! -d "yaml/excel_1_7" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/excel_1_7 --output-folder ./yaml/excel_1_7 --office 2>/dev/null
fi
if [ ! -d "yaml/excel_1_8" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/excel_1_8 --output-folder ./yaml/excel_1_8 --office 2>/dev/null
fi
if [ ! -d "yaml/excel_1_9" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/excel_1_9 --output-folder ./yaml/excel_1_9 --office 2>/dev/null
fi
if [ ! -d "yaml/excel_1_10" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/excel_1_10 --output-folder ./yaml/excel_1_10 --office 2>/dev/null
fi
if [ ! -d "yaml/excel_1_11" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/excel_1_11 --output-folder ./yaml/excel_1_11 --office 2>/dev/null
fi
if [ ! -d "yaml/excel_1_12" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/excel_1_12 --output-folder ./yaml/excel_1_12 --office 2>/dev/null
fi
if [ ! -d "yaml/excel_1_13" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/excel_1_13 --output-folder ./yaml/excel_1_13 --office 2>/dev/null
fi
if [ ! -d "yaml/excel_1_14" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/excel_1_14 --output-folder ./yaml/excel_1_14 --office 2>/dev/null
fi
if [ ! -d "yaml/excel_1_15" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/excel_1_15 --output-folder ./yaml/excel_1_15 --office 2>/dev/null
fi
if [ ! -d "yaml/excel_1_16" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/excel_1_16 --output-folder ./yaml/excel_1_16 --office 2>/dev/null
fi
if [ ! -d "yaml/excel_1_17" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/excel_1_17 --output-folder ./yaml/excel_1_17 --office 2>/dev/null
fi
if [ ! -d "yaml/excel_online" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/excel_online --output-folder ./yaml/excel_online --office 2>/dev/null
fi
if [ ! -d "yaml/onenote" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/onenote --output-folder ./yaml/onenote --office
fi
if [ ! -d "yaml/outlook" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/outlook --output-folder ./yaml/outlook --office
fi
if [ ! -d "yaml/outlook_1_1" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/outlook_1_1 --output-folder ./yaml/outlook_1_1 --office 2>/dev/null
fi
if [ ! -d "yaml/outlook_1_2" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/outlook_1_2 --output-folder ./yaml/outlook_1_2 --office 2>/dev/null
fi
if [ ! -d "yaml/outlook_1_3" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/outlook_1_3 --output-folder ./yaml/outlook_1_3 --office 2>/dev/null
fi
if [ ! -d "yaml/outlook_1_4" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/outlook_1_4 --output-folder ./yaml/outlook_1_4 --office 2>/dev/null
fi
if [ ! -d "yaml/outlook_1_5" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/outlook_1_5 --output-folder ./yaml/outlook_1_5 --office 2>/dev/null
fi
if [ ! -d "yaml/outlook_1_6" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/outlook_1_6 --output-folder ./yaml/outlook_1_6 --office 2>/dev/null
fi
if [ ! -d "yaml/outlook_1_7" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/outlook_1_7 --output-folder ./yaml/outlook_1_7 --office 2>/dev/null
fi
if [ ! -d "yaml/outlook_1_8" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/outlook_1_8 --output-folder ./yaml/outlook_1_8 --office 2>/dev/null
fi
if [ ! -d "yaml/outlook_1_9" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/outlook_1_9 --output-folder ./yaml/outlook_1_9 --office 2>/dev/null
fi
if [ ! -d "yaml/outlook_1_10" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/outlook_1_10 --output-folder ./yaml/outlook_1_10 --office 2>/dev/null
fi
if [ ! -d "yaml/outlook_1_11" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/outlook_1_11 --output-folder ./yaml/outlook_1_11 --office 2>/dev/null
fi
if [ ! -d "yaml/outlook_1_12" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/outlook_1_12 --output-folder ./yaml/outlook_1_12 --office 2>/dev/null
fi
if [ ! -d "yaml/outlook_1_13" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/outlook_1_13 --output-folder ./yaml/outlook_1_13 --office 2>/dev/null
fi
if [ ! -d "yaml/outlook_1_14" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/outlook_1_14 --output-folder ./yaml/outlook_1_14 --office 2>/dev/null
fi
if [ ! -d "yaml/powerpoint" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/powerpoint --output-folder ./yaml/powerpoint --office
fi
if [ ! -d "yaml/powerpoint_1_1" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/powerpoint_1_1 --output-folder ./yaml/powerpoint_1_1 --office 2>/dev/null
fi
if [ ! -d "yaml/powerpoint_1_2" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/powerpoint_1_2 --output-folder ./yaml/powerpoint_1_2 --office 2>/dev/null
fi
if [ ! -d "yaml/powerpoint_1_3" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/powerpoint_1_3 --output-folder ./yaml/powerpoint_1_3 --office 2>/dev/null
fi
if [ ! -d "yaml/powerpoint_1_4" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/powerpoint_1_4 --output-folder ./yaml/powerpoint_1_4 --office 2>/dev/null
fi
if [ ! -d "yaml/powerpoint_1_5" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/powerpoint_1_5 --output-folder ./yaml/powerpoint_1_5 --office 2>/dev/null
fi
if [ ! -d "yaml/powerpoint_1_6" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/powerpoint_1_6 --output-folder ./yaml/powerpoint_1_6 --office 2>/dev/null
fi
if [ ! -d "yaml/powerpoint_1_7" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/powerpoint_1_7 --output-folder ./yaml/powerpoint_1_7 --office 2>/dev/null
fi
if [ ! -d "yaml/visio" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/visio --output-folder ./yaml/visio --office
fi
if [ ! -d "yaml/word" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/word --output-folder ./yaml/word --office
fi
if [ ! -d "yaml/word_1_1" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/word_1_1 --output-folder ./yaml/word_1_1 --office 2>/dev/null
fi
if [ ! -d "yaml/word_1_2" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/word_1_2 --output-folder ./yaml/word_1_2 --office 2>/dev/null
fi
if [ ! -d "yaml/word_1_3" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/word_1_3 --output-folder ./yaml/word_1_3 --office 2>/dev/null
fi
if [ ! -d "yaml/word_1_4" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/word_1_4 --output-folder ./yaml/word_1_4 --office 2>/dev/null
fi
if [ ! -d "yaml/word_1_5" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/word_1_5 --output-folder ./yaml/word_1_5 --office 2>/dev/null
fi
if [ ! -d "yaml/word_1_6" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/word_1_6 --output-folder ./yaml/word_1_6 --office 2>/dev/null
fi
if [ ! -d "yaml/word_1_7" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/word_1_7 --output-folder ./yaml/word_1_7 --office 2>/dev/null
fi
if [ ! -d "yaml/word_1_8" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/word_1_8 --output-folder ./yaml/word_1_8 --office 2>/dev/null
fi
if [ ! -d "yaml/word_1_9" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/word_1_9 --output-folder ./yaml/word_1_9 --office 2>/dev/null
fi
if [ ! -d "yaml/word_1_3_hidden_document" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/word_1_3_hidden_document --output-folder ./yaml/word_1_3_hidden_document --office 2>/dev/null
fi
if [ ! -d "yaml/word_1_4_hidden_document" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/word_1_4_hidden_document --output-folder ./yaml/word_1_4_hidden_document --office 2>/dev/null
fi
if [ ! -d "yaml/word_1_5_hidden_document" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/word_1_5_hidden_document --output-folder ./yaml/word_1_5_hidden_document --office 2>/dev/null
fi
if [ ! -d "yaml/word_desktop_1_1" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/word_desktop_1_1 --output-folder ./yaml/word_desktop_1_1 --office 2>/dev/null
fi
if [ ! -d "yaml/word_online" ]; then
    ./node_modules/.bin/api-documenter yaml --input-folder ./json/word_online --output-folder ./yaml/word_online --office 2>/dev/null
fi

pushd scripts
node postprocessor.js
popd

./node_modules/.bin/reference-coverage-tester reference-coverage-tester.json
