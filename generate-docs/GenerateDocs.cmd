IF EXIST "node_modules" (
    rmdir "node_modules" /s /q
)

IF EXIST "scripts\node_modules" (
    rmdir "scripts\node_modules" /s /q
)

IF EXIST "tools\node_modules" (
    rmdir "tools\node_modules" /s /q
)

IF EXIST "json" (
    rmdir "json" /s /q
)

call md json

IF EXIST "yaml" (
    rmdir "yaml" /s /q
)

call md yaml

call npm install

pushd scripts
call npm install
call npm run build
call node preprocessor.js
popd


pushd tools
call npm install
call npm run build
call node version-remover ..\api-extractor-inputs-excel-release\Excel_online\excel.d.ts "ExcelApi Online" ..\api-extractor-inputs-excel-release\Excel_1_10\excel.d.ts
call node version-remover ..\api-extractor-inputs-excel-release\Excel_1_10\excel.d.ts "ExcelApi 1.10" ..\api-extractor-inputs-excel-release\Excel_1_9\excel.d.ts
call node version-remover ..\api-extractor-inputs-excel-release\Excel_1_9\excel.d.ts "ExcelApi 1.9" ..\api-extractor-inputs-excel-release\Excel_1_8\excel.d.ts
call node version-remover ..\api-extractor-inputs-excel-release\Excel_1_8\excel.d.ts "ExcelApi 1.8" ..\api-extractor-inputs-excel-release\Excel_1_7\excel.d.ts
call node version-remover ..\api-extractor-inputs-excel-release\Excel_1_7\excel.d.ts "ExcelApi 1.7" ..\api-extractor-inputs-excel-release\Excel_1_6\excel.d.ts
call node version-remover ..\api-extractor-inputs-excel-release\Excel_1_6\excel.d.ts "ExcelApi 1.6" ..\api-extractor-inputs-excel-release\Excel_1_5\excel.d.ts
call node version-remover ..\api-extractor-inputs-excel-release\Excel_1_5\excel.d.ts "ExcelApi 1.5" ..\api-extractor-inputs-excel-release\Excel_1_4\excel.d.ts
call node version-remover ..\api-extractor-inputs-excel-release\Excel_1_4\excel.d.ts "ExcelApi 1.4" ..\api-extractor-inputs-excel-release\Excel_1_3\excel.d.ts
call node version-remover ..\api-extractor-inputs-excel-release\Excel_1_3\excel.d.ts "ExcelApi 1.3" ..\api-extractor-inputs-excel-release\Excel_1_2\excel.d.ts
call node version-remover ..\api-extractor-inputs-excel-release\Excel_1_2\excel.d.ts "ExcelApi 1.2" ..\api-extractor-inputs-excel-release\Excel_1_1\excel.d.ts

call node version-remover ..\api-extractor-inputs-outlook-release\outlook_1_8\outlook.d.ts "Mailbox 1.8" ..\api-extractor-inputs-outlook-release\outlook_1_7\outlook.d.ts
call node version-remover ..\api-extractor-inputs-outlook-release\outlook_1_7\outlook.d.ts "Mailbox 1.7" ..\api-extractor-inputs-outlook-release\outlook_1_6\outlook.d.ts
call node version-remover ..\api-extractor-inputs-outlook-release\outlook_1_6\outlook.d.ts "Mailbox 1.6" ..\api-extractor-inputs-outlook-release\outlook_1_5\outlook.d.ts
call node version-remover ..\api-extractor-inputs-outlook-release\outlook_1_5\outlook.d.ts "Mailbox 1.5" ..\api-extractor-inputs-outlook-release\outlook_1_4\outlook.d.ts
call node version-remover ..\api-extractor-inputs-outlook-release\outlook_1_4\outlook.d.ts "Mailbox 1.4" ..\api-extractor-inputs-outlook-release\outlook_1_3\outlook.d.ts
call node version-remover ..\api-extractor-inputs-outlook-release\outlook_1_3\outlook.d.ts "Mailbox 1.3" ..\api-extractor-inputs-outlook-release\outlook_1_2\outlook.d.ts
call node version-remover ..\api-extractor-inputs-outlook-release\outlook_1_2\outlook.d.ts "Mailbox 1.2" ..\api-extractor-inputs-outlook-release\outlook_1_1\outlook.d.ts

call node version-remover ..\api-extractor-inputs-word-release\word_1_3\word.d.ts "WordApi 1.3" ..\api-extractor-inputs-word-release\word_1_2\word.d.ts
call node version-remover ..\api-extractor-inputs-word-release\word_1_2\word.d.ts "WordApi 1.2" ..\api-extractor-inputs-word-release\word_1_1\word.d.ts
popd

cd api-extractor-inputs-office
call ..\node_modules\.bin\api-extractor run

cd ..\api-extractor-inputs-excel
call ..\node_modules\.bin\api-extractor run
cd ..\api-extractor-inputs-excel-release\excel_online
call ..\..\node_modules\.bin\api-extractor run
cd ..\excel_1_10
call ..\..\node_modules\.bin\api-extractor run
cd ..\excel_1_9
call ..\..\node_modules\.bin\api-extractor run
cd ..\excel_1_8
call ..\..\node_modules\.bin\api-extractor run
cd ..\excel_1_7
call ..\..\node_modules\.bin\api-extractor run
cd ..\excel_1_6
call ..\..\node_modules\.bin\api-extractor run
cd ..\excel_1_5
call ..\..\node_modules\.bin\api-extractor run
cd ..\excel_1_4
call ..\..\node_modules\.bin\api-extractor run
cd ..\excel_1_3
call ..\..\node_modules\.bin\api-extractor run
cd ..\excel_1_2
call ..\..\node_modules\.bin\api-extractor run
cd ..\excel_1_1
call ..\..\node_modules\.bin\api-extractor run
cd ..

cd ..\api-extractor-inputs-onenote
call ..\node_modules\.bin\api-extractor run

cd ..\api-extractor-inputs-outlook
call ..\node_modules\.bin\api-extractor run
cd ..\api-extractor-inputs-outlook-release\outlook_1_8
call ..\..\node_modules\.bin\api-extractor run
cd ..\outlook_1_7
call node ..\..\scripts\versioned-dts-cleanup outlook.d.ts Outlook 1.7
call ..\..\node_modules\.bin\api-extractor run
cd ..\outlook_1_6
call node ..\..\scripts\versioned-dts-cleanup outlook.d.ts Outlook 1.6
call ..\..\node_modules\.bin\api-extractor run
cd ..\outlook_1_5
call node ..\..\scripts\versioned-dts-cleanup outlook.d.ts Outlook 1.5
call ..\..\node_modules\.bin\api-extractor run
cd ..\outlook_1_4
call node ..\..\scripts\versioned-dts-cleanup outlook.d.ts Outlook 1.4
call ..\..\node_modules\.bin\api-extractor run
cd ..\outlook_1_3
call node ..\..\scripts\versioned-dts-cleanup outlook.d.ts Outlook 1.3
call ..\..\node_modules\.bin\api-extractor run
cd ..\outlook_1_2
call node ..\..\scripts\versioned-dts-cleanup outlook.d.ts Outlook 1.2
call ..\..\node_modules\.bin\api-extractor run
cd ..\outlook_1_1
call node ..\..\scripts\versioned-dts-cleanup outlook.d.ts Outlook 1.1
call ..\..\node_modules\.bin\api-extractor run
cd ..

cd ..\api-extractor-inputs-powerpoint
call ..\node_modules\.bin\api-extractor run

cd ..\api-extractor-inputs-visio
call ..\node_modules\.bin\api-extractor run

cd ..\api-extractor-inputs-word
call ..\node_modules\.bin\api-extractor run
cd ..\api-extractor-inputs-word-release\word_1_3
call ..\..\node_modules\.bin\api-extractor run
cd ..\word_1_2
call ..\..\node_modules\.bin\api-extractor run
cd ..\word_1_1
call ..\..\node_modules\.bin\api-extractor run
cd ..

cd ..\api-extractor-inputs-custom-functions-runtime
call ..\node_modules\.bin\api-extractor run

cd ..\api-extractor-inputs-office-runtime
call ..\node_modules\.bin\api-extractor run

cd ..

pushd scripts
call node midprocessor.js
popd


call .\node_modules\.bin\api-documenter yaml --input-folder .\json\office --output-folder .\yaml\office --office

call .\node_modules\.bin\api-documenter yaml --input-folder .\json\excel --output-folder .\yaml\excel --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\excel_1_1 --output-folder .\yaml\excel_1_1 --office 2> nul
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\excel_1_2 --output-folder .\yaml\excel_1_2 --office 2> nul
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\excel_1_3 --output-folder .\yaml\excel_1_3 --office 2> nul
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\excel_1_4 --output-folder .\yaml\excel_1_4 --office 2> nul
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\excel_1_5 --output-folder .\yaml\excel_1_5 --office 2> nul
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\excel_1_6 --output-folder .\yaml\excel_1_6 --office 2> nul
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\excel_1_7 --output-folder .\yaml\excel_1_7 --office 2> nul
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\excel_1_8 --output-folder .\yaml\excel_1_8 --office 2> nul
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\excel_1_9 --output-folder .\yaml\excel_1_9 --office 2> nul
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\excel_1_10 --output-folder .\yaml\excel_1_10 --office 2> nul
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\excel_online --output-folder .\yaml\excel_online --office 2> nul
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\onenote --output-folder .\yaml\onenote --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\outlook --output-folder .\yaml\outlook --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\outlook_1_1 --output-folder .\yaml\outlook_1_1 --office 2> nul
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\outlook_1_2 --output-folder .\yaml\outlook_1_2 --office 2> nul
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\outlook_1_3 --output-folder .\yaml\outlook_1_3 --office 2> nul
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\outlook_1_4 --output-folder .\yaml\outlook_1_4 --office 2> nul
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\outlook_1_5 --output-folder .\yaml\outlook_1_5 --office 2> nul
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\outlook_1_6 --output-folder .\yaml\outlook_1_6 --office 2> nul
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\outlook_1_7 --output-folder .\yaml\outlook_1_7 --office 2> nul
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\outlook_1_8 --output-folder .\yaml\outlook_1_8 --office 2> nul
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\powerpoint --output-folder .\yaml\powerpoint --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\visio --output-folder .\yaml\visio --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\word --output-folder .\yaml\word --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\word_1_1 --output-folder .\yaml\word_1_1 --office 2> nul
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\word_1_2 --output-folder .\yaml\word_1_2 --office 2> nul
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\word_1_3 --output-folder .\yaml\word_1_3 --office 2> nul


pushd scripts
call node postprocessor.js
popd

pause