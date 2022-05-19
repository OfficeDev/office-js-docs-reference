IF EXIST "node_modules" (
    rmdir "node_modules" /s /q
)

IF EXIST "scripts\node_modules" (
    rmdir "scripts\node_modules" /s /q
)

IF EXIST "tools\node_modules" (
    rmdir "tools\node_modules" /s /q
)

IF NOT EXIST "json" (
    call md json
)

IF NOT EXIST "yaml" (
    call md yaml
)

call npm install

pushd scripts
call npm install
call npm run build
call node preprocessor.js
popd


pushd tools
call md tool-inputs
call npm install
call npm run build
call node version-remover ..\api-extractor-inputs-excel-release\Excel_online\excel.d.ts "ExcelApiOnline 1.1" ..\api-extractor-inputs-excel-release\Excel_1_15\excel.d.ts
call node version-remover ..\api-extractor-inputs-excel-release\Excel_1_14\excel.d.ts "ExcelApi 1.15" ..\api-extractor-inputs-excel-release\Excel_1_14\excel.d.ts
call node version-remover ..\api-extractor-inputs-excel-release\Excel_1_14\excel.d.ts "ExcelApi 1.14" ..\api-extractor-inputs-excel-release\Excel_1_13\excel.d.ts
call node version-remover ..\api-extractor-inputs-excel-release\Excel_1_13\excel.d.ts "ExcelApi 1.13" ..\api-extractor-inputs-excel-release\Excel_1_12\excel.d.ts
call node version-remover ..\api-extractor-inputs-excel-release\Excel_1_12\excel.d.ts "ExcelApi 1.12" ..\api-extractor-inputs-excel-release\Excel_1_11\excel.d.ts
call node version-remover ..\api-extractor-inputs-excel-release\Excel_1_11\excel.d.ts "ExcelApi 1.11" ..\api-extractor-inputs-excel-release\Excel_1_10\excel.d.ts
call node version-remover ..\api-extractor-inputs-excel-release\Excel_1_10\excel.d.ts "ExcelApi 1.10" ..\api-extractor-inputs-excel-release\Excel_1_9\excel.d.ts
call node version-remover ..\api-extractor-inputs-excel-release\Excel_1_9\excel.d.ts "ExcelApi 1.9" ..\api-extractor-inputs-excel-release\Excel_1_8\excel.d.ts
call node version-remover ..\api-extractor-inputs-excel-release\Excel_1_8\excel.d.ts "ExcelApi 1.8" ..\api-extractor-inputs-excel-release\Excel_1_7\excel.d.ts
call node version-remover ..\api-extractor-inputs-excel-release\Excel_1_7\excel.d.ts "ExcelApi 1.7" ..\api-extractor-inputs-excel-release\Excel_1_6\excel.d.ts
call node version-remover ..\api-extractor-inputs-excel-release\Excel_1_6\excel.d.ts "ExcelApi 1.6" ..\api-extractor-inputs-excel-release\Excel_1_5\excel.d.ts
call node version-remover ..\api-extractor-inputs-excel-release\Excel_1_5\excel.d.ts "ExcelApi 1.5" ..\api-extractor-inputs-excel-release\Excel_1_4\excel.d.ts
call node version-remover ..\api-extractor-inputs-excel-release\Excel_1_4\excel.d.ts "ExcelApi 1.4" ..\api-extractor-inputs-excel-release\Excel_1_3\excel.d.ts
call node version-remover ..\api-extractor-inputs-excel-release\Excel_1_3\excel.d.ts "ExcelApi 1.3" ..\api-extractor-inputs-excel-release\Excel_1_2\excel.d.ts
call node version-remover ..\api-extractor-inputs-excel-release\Excel_1_2\excel.d.ts "ExcelApi 1.2" ..\api-extractor-inputs-excel-release\Excel_1_1\excel.d.ts
call node version-remover ..\api-extractor-inputs-excel-release\Excel_1_1\excel.d.ts "ExcelApi 1.1" .\tool-inputs\excel-base.d.ts

call node version-remover ..\api-extractor-inputs-outlook-release\outlook_1_11\outlook.d.ts "Mailbox 1.11" ..\api-extractor-inputs-outlook-release\outlook_1_10\outlook.d.ts
call node ..\scripts\versioned-dts-cleanup ..\api-extractor-inputs-outlook-release\outlook_1_10\outlook.d.ts Outlook 1.10
call node version-remover ..\api-extractor-inputs-outlook-release\outlook_1_10\outlook.d.ts "Mailbox 1.10" ..\api-extractor-inputs-outlook-release\outlook_1_9\outlook.d.ts
call node ..\scripts\versioned-dts-cleanup ..\api-extractor-inputs-outlook-release\outlook_1_9\outlook.d.ts Outlook 1.9
call node version-remover ..\api-extractor-inputs-outlook-release\outlook_1_9\outlook.d.ts "Mailbox 1.9" ..\api-extractor-inputs-outlook-release\outlook_1_8\outlook.d.ts
call node ..\scripts\versioned-dts-cleanup ..\api-extractor-inputs-outlook-release\outlook_1_8\outlook.d.ts Outlook 1.8
call node version-remover ..\api-extractor-inputs-outlook-release\outlook_1_8\outlook.d.ts "Mailbox 1.8" ..\api-extractor-inputs-outlook-release\outlook_1_7\outlook.d.ts
call node ..\scripts\versioned-dts-cleanup ..\api-extractor-inputs-outlook-release\outlook_1_7\outlook.d.ts Outlook 1.7
call node version-remover ..\api-extractor-inputs-outlook-release\outlook_1_7\outlook.d.ts "Mailbox 1.7" ..\api-extractor-inputs-outlook-release\outlook_1_6\outlook.d.ts
call node ..\scripts\versioned-dts-cleanup ..\api-extractor-inputs-outlook-release\outlook_1_6\outlook.d.ts Outlook 1.6
call node version-remover ..\api-extractor-inputs-outlook-release\outlook_1_6\outlook.d.ts "Mailbox 1.6" ..\api-extractor-inputs-outlook-release\outlook_1_5\outlook.d.ts
call node ..\scripts\versioned-dts-cleanup ..\api-extractor-inputs-outlook-release\outlook_1_5\outlook.d.ts Outlook 1.5
call node version-remover ..\api-extractor-inputs-outlook-release\outlook_1_5\outlook.d.ts "Mailbox 1.5" ..\api-extractor-inputs-outlook-release\outlook_1_4\outlook.d.ts
call node ..\scripts\versioned-dts-cleanup ..\api-extractor-inputs-outlook-release\outlook_1_4\outlook.d.ts Outlook 1.4
call node version-remover ..\api-extractor-inputs-outlook-release\outlook_1_4\outlook.d.ts "Mailbox 1.4" ..\api-extractor-inputs-outlook-release\outlook_1_3\outlook.d.ts
call node ..\scripts\versioned-dts-cleanup ..\api-extractor-inputs-outlook-release\outlook_1_3\outlook.d.ts Outlook 1.3
call node version-remover ..\api-extractor-inputs-outlook-release\outlook_1_3\outlook.d.ts "Mailbox 1.3" ..\api-extractor-inputs-outlook-release\outlook_1_2\outlook.d.ts
call node ..\scripts\versioned-dts-cleanup ..\api-extractor-inputs-outlook-release\outlook_1_2\outlook.d.ts Outlook 1.2
call node version-remover ..\api-extractor-inputs-outlook-release\outlook_1_2\outlook.d.ts "Mailbox 1.2" ..\api-extractor-inputs-outlook-release\outlook_1_1\outlook.d.ts
call node ..\scripts\versioned-dts-cleanup ..\api-extractor-inputs-outlook-release\outlook_1_1\outlook.d.ts Outlook 1.1
call node version-remover ..\api-extractor-inputs-outlook-release\outlook_1_1\outlook.d.ts "Mailbox 1.1" .\tool-inputs\outlook-base.d.ts

call node version-remover ..\api-extractor-inputs-powerpoint-release\powerpoint_1_3\powerpoint.d.ts "PowerPointApi 1.3" ..\api-extractor-inputs-powerpoint-release\powerpoint_1_2\powerpoint.d.ts
call node version-remover ..\api-extractor-inputs-powerpoint-release\powerpoint_1_2\powerpoint.d.ts "PowerPointApi 1.2" ..\api-extractor-inputs-powerpoint-release\powerpoint_1_1\powerpoint.d.ts
call node version-remover ..\api-extractor-inputs-powerpoint-release\powerpoint_1_1\powerpoint.d.ts "PowerPointApi 1.1" .\tool-inputs\powerpoint-base.d.ts


call node version-remover ..\api-extractor-inputs-word-release\word_1_3\word.d.ts "WordApi 1.3" ..\api-extractor-inputs-word-release\word_1_2\word.d.ts
call node version-remover ..\api-extractor-inputs-word-release\word_1_2\word.d.ts "WordApi 1.2" ..\api-extractor-inputs-word-release\word_1_1\word.d.ts
call node version-remover ..\api-extractor-inputs-word-release\word_1_1\word.d.ts "WordApi 1.1" .\tool-inputs\word-base.d.ts


call node whats-new excel ..\api-extractor-inputs-excel\excel.d.ts ..\api-extractor-inputs-excel-release\Excel_online\excel.d.ts ..\..\docs\includes\excel-preview
call node whats-new excel ..\api-extractor-inputs-excel-release\Excel_online\excel.d.ts ..\api-extractor-inputs-excel-release\Excel_1_15\excel.d.ts ..\..\docs\includes\excel-online
call node whats-new excel ..\api-extractor-inputs-excel-release\Excel_1_14\excel.d.ts ..\api-extractor-inputs-excel-release\Excel_1_14\excel.d.ts ..\..\docs\includes\excel-1_15
call node whats-new excel ..\api-extractor-inputs-excel-release\Excel_1_14\excel.d.ts ..\api-extractor-inputs-excel-release\Excel_1_13\excel.d.ts ..\..\docs\includes\excel-1_14
call node whats-new excel ..\api-extractor-inputs-excel-release\Excel_1_13\excel.d.ts ..\api-extractor-inputs-excel-release\Excel_1_12\excel.d.ts ..\..\docs\includes\excel-1_13
call node whats-new excel ..\api-extractor-inputs-excel-release\Excel_1_12\excel.d.ts ..\api-extractor-inputs-excel-release\Excel_1_11\excel.d.ts ..\..\docs\includes\excel-1_12
call node whats-new excel ..\api-extractor-inputs-excel-release\Excel_1_11\excel.d.ts ..\api-extractor-inputs-excel-release\Excel_1_10\excel.d.ts ..\..\docs\includes\excel-1_11
call node whats-new excel ..\api-extractor-inputs-excel-release\Excel_1_10\excel.d.ts ..\api-extractor-inputs-excel-release\Excel_1_9\excel.d.ts ..\..\docs\includes\excel-1_10
call node whats-new excel ..\api-extractor-inputs-excel-release\Excel_1_9\excel.d.ts ..\api-extractor-inputs-excel-release\Excel_1_8\excel.d.ts ..\..\docs\includes\excel-1_9
call node whats-new excel ..\api-extractor-inputs-excel-release\Excel_1_8\excel.d.ts ..\api-extractor-inputs-excel-release\Excel_1_7\excel.d.ts ..\..\docs\includes\excel-1_8
call node whats-new excel ..\api-extractor-inputs-excel-release\Excel_1_7\excel.d.ts ..\api-extractor-inputs-excel-release\Excel_1_6\excel.d.ts ..\..\docs\includes\excel-1_7
call node whats-new excel ..\api-extractor-inputs-excel-release\Excel_1_6\excel.d.ts ..\api-extractor-inputs-excel-release\Excel_1_5\excel.d.ts ..\..\docs\includes\excel-1_6
call node whats-new excel ..\api-extractor-inputs-excel-release\Excel_1_5\excel.d.ts ..\api-extractor-inputs-excel-release\Excel_1_4\excel.d.ts ..\..\docs\includes\excel-1_5
call node whats-new excel ..\api-extractor-inputs-excel-release\Excel_1_4\excel.d.ts ..\api-extractor-inputs-excel-release\Excel_1_3\excel.d.ts ..\..\docs\includes\excel-1_4
call node whats-new excel ..\api-extractor-inputs-excel-release\Excel_1_3\excel.d.ts ..\api-extractor-inputs-excel-release\Excel_1_2\excel.d.ts ..\..\docs\includes\excel-1_3
call node whats-new excel ..\api-extractor-inputs-excel-release\Excel_1_2\excel.d.ts ..\api-extractor-inputs-excel-release\Excel_1_1\excel.d.ts ..\..\docs\includes\excel-1_2
call node whats-new excel ..\api-extractor-inputs-excel-release\Excel_1_1\excel.d.ts .\tool-inputs\excel-base.d.ts ..\..\docs\includes\excel-1_1

call node whats-new outlook ..\api-extractor-inputs-outlook\outlook.d.ts ..\api-extractor-inputs-outlook-release\outlook_1_11\outlook.d.ts ..\..\docs\includes\outlook-preview
call node whats-new outlook ..\api-extractor-inputs-outlook-release\outlook_1_11\outlook.d.ts ..\api-extractor-inputs-outlook-release\outlook_1_10\outlook.d.ts ..\..\docs\includes\outlook-1_11
call node whats-new outlook ..\api-extractor-inputs-outlook-release\outlook_1_10\outlook.d.ts ..\api-extractor-inputs-outlook-release\outlook_1_9\outlook.d.ts ..\..\docs\includes\outlook-1_10
call node whats-new outlook ..\api-extractor-inputs-outlook-release\outlook_1_9\outlook.d.ts ..\api-extractor-inputs-outlook-release\outlook_1_8\outlook.d.ts ..\..\docs\includes\outlook-1_9
call node whats-new outlook ..\api-extractor-inputs-outlook-release\outlook_1_8\outlook.d.ts ..\api-extractor-inputs-outlook-release\outlook_1_7\outlook.d.ts ..\..\docs\includes\outlook-1_8
call node whats-new outlook ..\api-extractor-inputs-outlook-release\outlook_1_7\outlook.d.ts ..\api-extractor-inputs-outlook-release\outlook_1_6\outlook.d.ts ..\..\docs\includes\outlook-1_7
call node whats-new outlook ..\api-extractor-inputs-outlook-release\outlook_1_6\outlook.d.ts ..\api-extractor-inputs-outlook-release\outlook_1_5\outlook.d.ts ..\..\docs\includes\outlook-1_6
call node whats-new outlook ..\api-extractor-inputs-outlook-release\outlook_1_5\outlook.d.ts ..\api-extractor-inputs-outlook-release\outlook_1_4\outlook.d.ts ..\..\docs\includes\outlook-1_5
call node whats-new outlook ..\api-extractor-inputs-outlook-release\outlook_1_4\outlook.d.ts ..\api-extractor-inputs-outlook-release\outlook_1_3\outlook.d.ts ..\..\docs\includes\outlook-1_4
call node whats-new outlook ..\api-extractor-inputs-outlook-release\outlook_1_3\outlook.d.ts ..\api-extractor-inputs-outlook-release\outlook_1_2\outlook.d.ts ..\..\docs\includes\outlook-1_3
call node whats-new outlook ..\api-extractor-inputs-outlook-release\outlook_1_2\outlook.d.ts ..\api-extractor-inputs-outlook-release\outlook_1_1\outlook.d.ts ..\..\docs\includes\outlook-1_2
call node whats-new outlook ..\api-extractor-inputs-outlook-release\outlook_1_1\outlook.d.ts .\tool-inputs\outlook-base.d.ts ..\..\docs\includes\outlook-1_1

call node whats-new powerpoint ..\api-extractor-inputs-powerpoint\powerpoint.d.ts ..\api-extractor-inputs-powerpoint-release\powerpoint_1_3\powerpoint.d.ts ..\..\docs\includes\powerpoint-preview
call node whats-new powerpoint ..\api-extractor-inputs-powerpoint-release\powerpoint_1_3\powerpoint.d.ts ..\api-extractor-inputs-powerpoint-release\powerpoint_1_2\powerpoint.d.ts ..\..\docs\includes\powerpoint-1_3
call node whats-new powerpoint ..\api-extractor-inputs-powerpoint-release\powerpoint_1_2\powerpoint.d.ts ..\api-extractor-inputs-powerpoint-release\powerpoint_1_1\powerpoint.d.ts ..\..\docs\includes\powerpoint-1_2
call node whats-new powerpoint ..\api-extractor-inputs-powerpoint-release\powerpoint_1_1\powerpoint.d.ts .\tool-inputs\powerpoint-base.d.ts ..\..\docs\includes\powerpoint-1_1

call node whats-new word ..\api-extractor-inputs-word\word.d.ts ..\api-extractor-inputs-word-release\word_1_3\word.d.ts ..\..\docs\includes\word-preview
call node whats-new word ..\api-extractor-inputs-word-release\word_1_3\word.d.ts ..\api-extractor-inputs-word-release\word_1_2\word.d.ts ..\..\docs\includes\word-1_3
call node whats-new word ..\api-extractor-inputs-word-release\word_1_2\word.d.ts ..\api-extractor-inputs-word-release\word_1_1\word.d.ts ..\..\docs\includes\word-1_2
call node whats-new word ..\api-extractor-inputs-word-release\word_1_1\word.d.ts .\tool-inputs\word-base.d.ts ..\..\docs\includes\word-1_1

popd

if NOT EXIST "json/office" (
    echo Running API Extractor for Office preview.
    pushd api-extractor-inputs-office
    call ..\node_modules\.bin\api-extractor run
    popd
)
if NOT EXIST "json/office-release" (
    echo Running API Extractor for Office release.
    pushd api-extractor-inputs-office-release
    call ..\node_modules\.bin\api-extractor run
    popd
)

if NOT EXIST "json/excel" (
    echo Running API Extractor for Excel preview.
    pushd api-extractor-inputs-excel
    call ..\node_modules\.bin\api-extractor run
    popd
)
if NOT EXIST "json/excel_online" (
    echo Running API Extractor for Excel online.
    pushd api-extractor-inputs-excel-release\excel_online
    call ..\..\node_modules\.bin\api-extractor run
    popd
)
if NOT EXIST "json/excel_1_15" (
    echo Running API Extractor for Excel 1.15.
    pushd api-extractor-inputs-excel-release\excel_1_15
    call ..\..\node_modules\.bin\api-extractor run
    popd
)
if NOT EXIST "json/excel_1_14" (
    echo Running API Extractor for Excel 1.14.
    pushd api-extractor-inputs-excel-release\excel_1_14
    call ..\..\node_modules\.bin\api-extractor run
    popd
)
if NOT EXIST "json/excel_1_13" (
    echo Running API Extractor for Excel 1.13.
    pushd api-extractor-inputs-excel-release\excel_1_13
    call ..\..\node_modules\.bin\api-extractor run
    popd
)
if NOT EXIST "json/excel_1_12" (
    echo Running API Extractor for Excel 1.12.
    pushd api-extractor-inputs-excel-release\excel_1_12
    call ..\..\node_modules\.bin\api-extractor run
    popd
)
if NOT EXIST "json/excel_1_11" (
    echo Running API Extractor for Excel 1.11.
    pushd api-extractor-inputs-excel-release\excel_1_11
    call ..\..\node_modules\.bin\api-extractor run
    popd
)
if NOT EXIST "json/excel_1_10" (
    echo Running API Extractor for Excel 1.10.
    pushd api-extractor-inputs-excel-release\excel_1_10
    call ..\..\node_modules\.bin\api-extractor run
    popd
)
if NOT EXIST "json/excel_1_9" (
    echo Running API Extractor for Excel 1.9.
    pushd api-extractor-inputs-excel-release\excel_1_9
    call ..\..\node_modules\.bin\api-extractor run
    popd
)
if NOT EXIST "json/excel_1_8" (
    echo Running API Extractor for Excel 1.8.
    pushd api-extractor-inputs-excel-release\excel_1_8
    call ..\..\node_modules\.bin\api-extractor run
    popd
)
if NOT EXIST "json/excel_1_7" (
    echo Running API Extractor for Excel 1.7.
    pushd api-extractor-inputs-excel-release\excel_1_7
    call ..\..\node_modules\.bin\api-extractor run
    popd
)
if NOT EXIST "json/excel_1_6" (
    echo Running API Extractor for Excel 1.6.
    pushd api-extractor-inputs-excel-release\excel_1_6
    call ..\..\node_modules\.bin\api-extractor run
    popd
)
if NOT EXIST "json/excel_1_5" (
    echo Running API Extractor for Excel 1.5.
    pushd api-extractor-inputs-excel-release\excel_1_5
    call ..\..\node_modules\.bin\api-extractor run
    popd
)
if NOT EXIST "json/excel_1_4" (
    echo Running API Extractor for Excel 1.4.
    pushd api-extractor-inputs-excel-release\excel_1_4
    call ..\..\node_modules\.bin\api-extractor run
    popd
)
if NOT EXIST "json/excel_1_3" (
    echo Running API Extractor for Excel 1.3.
    pushd api-extractor-inputs-excel-release\excel_1_3
    call ..\..\node_modules\.bin\api-extractor run
    popd
)
if NOT EXIST "json/excel_1_2" (
    echo Running API Extractor for Excel 1.2.
    pushd api-extractor-inputs-excel-release\excel_1_2
    call ..\..\node_modules\.bin\api-extractor run
    popd
)
if NOT EXIST "json/excel_1_1" (
    echo Running API Extractor for Excel 1.1.
    pushd api-extractor-inputs-excel-release\excel_1_1
    call ..\..\node_modules\.bin\api-extractor run
    popd
)

if NOT EXIST "json/onenote" (
    echo Running API Extractor for OneNote.
    pushd api-extractor-inputs-onenote
    call ..\node_modules\.bin\api-extractor run
    popd
)

if NOT EXIST "json/outlook" (
    echo Running API Extractor for Outlook preview.
    pushd api-extractor-inputs-outlook
    call ..\node_modules\.bin\api-extractor run
    popd
)
if NOT EXIST "json/outlook_1_11" (
    echo Running API Extractor for Outlook 1.11.
    pushd api-extractor-inputs-outlook-release\outlook_1_11
    call ..\..\node_modules\.bin\api-extractor run
    popd
)
if NOT EXIST "json/outlook_1_10" (
    echo Running API Extractor for Outlook 1.10.
    pushd api-extractor-inputs-outlook-release\outlook_1_10
    call ..\..\node_modules\.bin\api-extractor run
    popd
)
if NOT EXIST "json/outlook_1_9" (
    echo Running API Extractor for Outlook 1.9.
    pushd api-extractor-inputs-outlook-release\outlook_1_9
    call ..\..\node_modules\.bin\api-extractor run
    popd
)
if NOT EXIST "json/outlook_1_8" (
    echo Running API Extractor for Outlook 1.8.
    pushd api-extractor-inputs-outlook-release\outlook_1_8
    call ..\..\node_modules\.bin\api-extractor run
    popd
)
if NOT EXIST "json/outlook_1_7" (
    echo Running API Extractor for Outlook 1.7.
    pushd api-extractor-inputs-outlook-release\outlook_1_7
    call ..\..\node_modules\.bin\api-extractor run
    popd
)
if NOT EXIST "json/outlook_1_6" (
    echo Running API Extractor for Outlook 1.6.
    pushd api-extractor-inputs-outlook-release\outlook_1_6
    call ..\..\node_modules\.bin\api-extractor run
    popd
)
if NOT EXIST "json/outlook_1_5" (
    echo Running API Extractor for Outlook 1.5.
    pushd api-extractor-inputs-outlook-release\outlook_1_5
    call ..\..\node_modules\.bin\api-extractor run
    popd
)
if NOT EXIST "json/outlook_1_4" (
    echo Running API Extractor for Outlook 1.4.
    pushd api-extractor-inputs-outlook-release\outlook_1_4
    call ..\..\node_modules\.bin\api-extractor run
    popd
)
if NOT EXIST "json/outlook_1_3" (
    echo Running API Extractor for Outlook 1.3.
    pushd api-extractor-inputs-outlook-release\outlook_1_3
    call ..\..\node_modules\.bin\api-extractor run
    popd
)
if NOT EXIST "json/outlook_1_2" (
    echo Running API Extractor for Outlook 1.2.
    pushd api-extractor-inputs-outlook-release\outlook_1_2
    call ..\..\node_modules\.bin\api-extractor run
    popd
)
if NOT EXIST "json/outlook_1_1" (
    echo Running API Extractor for Outlook 1.1.
    pushd api-extractor-inputs-outlook-release\outlook_1_1
    call ..\..\node_modules\.bin\api-extractor run
    popd
)

if NOT EXIST "json/powerpoint" (
    echo Running API Extractor for PowerPoint preview.
    pushd api-extractor-inputs-powerpoint
    call ..\node_modules\.bin\api-extractor run
    popd
)
if NOT EXIST "json/powerpoint_1_3" (
    echo Running API Extractor for PowerPoint 1.3.
    pushd api-extractor-inputs-powerpoint-release\PowerPoint_1_3
    call ..\..\node_modules\.bin\api-extractor run
    popd
)
if NOT EXIST "json/powerpoint_1_2" (
    echo Running API Extractor for PowerPoint 1.2.
    pushd api-extractor-inputs-powerpoint-release\PowerPoint_1_2
    call ..\..\node_modules\.bin\api-extractor run
    popd
)
if NOT EXIST "json/powerpoint_1_1" (
    echo Running API Extractor for PowerPoint 1.1.
    pushd api-extractor-inputs-powerpoint-release\PowerPoint_1_1
    call ..\..\node_modules\.bin\api-extractor run
    popd
)

if NOT EXIST "json/visio" (
    echo Running API Extractor for Visio.
    pushd api-extractor-inputs-visio
    call ..\node_modules\.bin\api-extractor run
    popd
)

if NOT EXIST "json/word" (
    echo Running API Extractor for Word preview.
    pushd api-extractor-inputs-word
    call ..\node_modules\.bin\api-extractor run
    popd
)
if NOT EXIST "json/word_1_3" (
    echo Running API Extractor for Word 1.3.
    pushd api-extractor-inputs-word-release\word_1_3
    call ..\..\node_modules\.bin\api-extractor run
    popd
)
if NOT EXIST "json/word_1_2" (
    echo Running API Extractor for Word 1.2.
    pushd api-extractor-inputs-word-release\word_1_2
    call ..\..\node_modules\.bin\api-extractor run
    popd
)
if NOT EXIST "json/word_1_1" (
    echo Running API Extractor for Word 1.1.
    pushd api-extractor-inputs-word-release\word_1_1
    call ..\..\node_modules\.bin\api-extractor run
    popd
)

if NOT EXIST "json/custom-functions-runtime.api.json" (
    echo Running API Extractor for Custom Functions.
    pushd api-extractor-inputs-custom-functions-runtime
    call ..\node_modules\.bin\api-extractor run
    popd
)

if NOT EXIST "json/office-runtime.api.json" (
    echo Running API Extractor for Office Runtime.
    pushd api-extractor-inputs-office-runtime
    call ..\node_modules\.bin\api-extractor run
    popd
)

pushd scripts
call node midprocessor.js
popd


if NOT EXIST "yaml/office" ( call .\node_modules\.bin\api-documenter yaml --input-folder .\json\office --output-folder .\yaml\office --office )
if NOT EXIST "yaml/office_release" ( call .\node_modules\.bin\api-documenter yaml --input-folder .\json\office_release --output-folder .\yaml\office_release --office 2> nul )

if NOT EXIST "yaml/excel" ( call .\node_modules\.bin\api-documenter yaml --input-folder .\json\excel --output-folder .\yaml\excel --office )
if NOT EXIST "yaml/excel_1_1" ( call .\node_modules\.bin\api-documenter yaml --input-folder .\json\excel_1_1 --output-folder .\yaml\excel_1_1 --office 2> nul )
if NOT EXIST "yaml/excel_1_2" ( call .\node_modules\.bin\api-documenter yaml --input-folder .\json\excel_1_2 --output-folder .\yaml\excel_1_2 --office 2> nul )
if NOT EXIST "yaml/excel_1_3" ( call .\node_modules\.bin\api-documenter yaml --input-folder .\json\excel_1_3 --output-folder .\yaml\excel_1_3 --office 2> nul )
if NOT EXIST "yaml/excel_1_4" ( call .\node_modules\.bin\api-documenter yaml --input-folder .\json\excel_1_4 --output-folder .\yaml\excel_1_4 --office 2> nul )
if NOT EXIST "yaml/excel_1_5" ( call .\node_modules\.bin\api-documenter yaml --input-folder .\json\excel_1_5 --output-folder .\yaml\excel_1_5 --office 2> nul )
if NOT EXIST "yaml/excel_1_6" ( call .\node_modules\.bin\api-documenter yaml --input-folder .\json\excel_1_6 --output-folder .\yaml\excel_1_6 --office 2> nul )
if NOT EXIST "yaml/excel_1_7" ( call .\node_modules\.bin\api-documenter yaml --input-folder .\json\excel_1_7 --output-folder .\yaml\excel_1_7 --office 2> nul )
if NOT EXIST "yaml/excel_1_8" ( call .\node_modules\.bin\api-documenter yaml --input-folder .\json\excel_1_8 --output-folder .\yaml\excel_1_8 --office 2> nul )
if NOT EXIST "yaml/excel_1_9" ( call .\node_modules\.bin\api-documenter yaml --input-folder .\json\excel_1_9 --output-folder .\yaml\excel_1_9 --office 2> nul )
if NOT EXIST "yaml/excel_1_10" ( call .\node_modules\.bin\api-documenter yaml --input-folder .\json\excel_1_10 --output-folder .\yaml\excel_1_10 --office 2> nul )
if NOT EXIST "yaml/excel_1_11" ( call .\node_modules\.bin\api-documenter yaml --input-folder .\json\excel_1_11 --output-folder .\yaml\excel_1_11 --office 2> nul )
if NOT EXIST "yaml/excel_1_12" ( call .\node_modules\.bin\api-documenter yaml --input-folder .\json\excel_1_12 --output-folder .\yaml\excel_1_12 --office 2> nul )
if NOT EXIST "yaml/excel_1_13" ( call .\node_modules\.bin\api-documenter yaml --input-folder .\json\excel_1_13 --output-folder .\yaml\excel_1_13 --office 2> nul )
if NOT EXIST "yaml/excel_1_14" ( call .\node_modules\.bin\api-documenter yaml --input-folder .\json\excel_1_14 --output-folder .\yaml\excel_1_14 --office 2> nul )
if NOT EXIST "yaml/excel_1_15" ( call .\node_modules\.bin\api-documenter yaml --input-folder .\json\excel_1_15 --output-folder .\yaml\excel_1_15 --office 2> nul )
if NOT EXIST "yaml/excel_online" ( call .\node_modules\.bin\api-documenter yaml --input-folder .\json\excel_online --output-folder .\yaml\excel_online --office 2> nul )
if NOT EXIST "yaml/onenote" ( call .\node_modules\.bin\api-documenter yaml --input-folder .\json\onenote --output-folder .\yaml\onenote --office )
if NOT EXIST "yaml/outlook" ( call .\node_modules\.bin\api-documenter yaml --input-folder .\json\outlook --output-folder .\yaml\outlook --office )
if NOT EXIST "yaml/outlook_1_1" ( call .\node_modules\.bin\api-documenter yaml --input-folder .\json\outlook_1_1 --output-folder .\yaml\outlook_1_1 --office 2> nul )
if NOT EXIST "yaml/outlook_1_2" ( call .\node_modules\.bin\api-documenter yaml --input-folder .\json\outlook_1_2 --output-folder .\yaml\outlook_1_2 --office 2> nul )
if NOT EXIST "yaml/outlook_1_3" ( call .\node_modules\.bin\api-documenter yaml --input-folder .\json\outlook_1_3 --output-folder .\yaml\outlook_1_3 --office 2> nul )
if NOT EXIST "yaml/outlook_1_4" ( call .\node_modules\.bin\api-documenter yaml --input-folder .\json\outlook_1_4 --output-folder .\yaml\outlook_1_4 --office 2> nul )
if NOT EXIST "yaml/outlook_1_5" ( call .\node_modules\.bin\api-documenter yaml --input-folder .\json\outlook_1_5 --output-folder .\yaml\outlook_1_5 --office 2> nul )
if NOT EXIST "yaml/outlook_1_6" ( call .\node_modules\.bin\api-documenter yaml --input-folder .\json\outlook_1_6 --output-folder .\yaml\outlook_1_6 --office 2> nul )
if NOT EXIST "yaml/outlook_1_7" ( call .\node_modules\.bin\api-documenter yaml --input-folder .\json\outlook_1_7 --output-folder .\yaml\outlook_1_7 --office 2> nul )
if NOT EXIST "yaml/outlook_1_8" ( call .\node_modules\.bin\api-documenter yaml --input-folder .\json\outlook_1_8 --output-folder .\yaml\outlook_1_8 --office 2> nul )
if NOT EXIST "yaml/outlook_1_9" ( call .\node_modules\.bin\api-documenter yaml --input-folder .\json\outlook_1_9 --output-folder .\yaml\outlook_1_9 --office 2> nul )
if NOT EXIST "yaml/outlook_1_10" ( call .\node_modules\.bin\api-documenter yaml --input-folder .\json\outlook_1_10 --output-folder .\yaml\outlook_1_10 --office 2> nul )
if NOT EXIST "yaml/outlook_1_11" ( call .\node_modules\.bin\api-documenter yaml --input-folder .\json\outlook_1_11 --output-folder .\yaml\outlook_1_11 --office 2> nul )
if NOT EXIST "yaml/powerpoint" ( call .\node_modules\.bin\api-documenter yaml --input-folder .\json\powerpoint --output-folder .\yaml\powerpoint --office )
if NOT EXIST "yaml/powerpoint_1_1" ( call .\node_modules\.bin\api-documenter yaml --input-folder .\json\powerpoint_1_1 --output-folder .\yaml\powerpoint_1_1 --office 2> nul )
if NOT EXIST "yaml/powerpoint_1_2" ( call .\node_modules\.bin\api-documenter yaml --input-folder .\json\powerpoint_1_2 --output-folder .\yaml\powerpoint_1_2 --office 2> nul )
if NOT EXIST "yaml/powerpoint_1_3" ( call .\node_modules\.bin\api-documenter yaml --input-folder .\json\powerpoint_1_3 --output-folder .\yaml\powerpoint_1_3 --office 2> nul )
if NOT EXIST "yaml/visio" ( call .\node_modules\.bin\api-documenter yaml --input-folder .\json\visio --output-folder .\yaml\visio --office )
if NOT EXIST "yaml/word" ( call .\node_modules\.bin\api-documenter yaml --input-folder .\json\word --output-folder .\yaml\word --office )
if NOT EXIST "yaml/word_1_1" ( call .\node_modules\.bin\api-documenter yaml --input-folder .\json\word_1_1 --output-folder .\yaml\word_1_1 --office 2> nul )
if NOT EXIST "yaml/word_1_2" ( call .\node_modules\.bin\api-documenter yaml --input-folder .\json\word_1_2 --output-folder .\yaml\word_1_2 --office 2> nul )
if NOT EXIST "yaml/word_1_3" ( call .\node_modules\.bin\api-documenter yaml --input-folder .\json\word_1_3 --output-folder .\yaml\word_1_3 --office 2> nul )

pushd scripts
call node postprocessor.js
popd

pushd tools
call node coverage-tester.js
popd

pause
