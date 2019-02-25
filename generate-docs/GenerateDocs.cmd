IF EXIST "node_modules" (
    rmdir "node_modules" /s /q
)

IF EXIST "scripts\node_modules" (
    rmdir "scripts\node_modules" /s /q
)

call npm install

pushd scripts
call npm install
call npm run build
REM call node preprocessor.js
popd

del package-lock.json

REM cd api-extractor-inputs-office
REM call ..\node_modules\.bin\api-extractor run

REM cd ..\api-extractor-inputs-excel
REM call ..\node_modules\.bin\api-extractor run
REM cd ..\api-extractor-inputs-excel-release\Excel_1_8
REM call ..\..\node_modules\.bin\api-extractor run
REM cd ..\Excel_1_7
REM call ..\..\node_modules\.bin\api-extractor run
REM cd ..\Excel_1_6
REM call ..\..\node_modules\.bin\api-extractor run
REM cd ..\Excel_1_5
REM call ..\..\node_modules\.bin\api-extractor run
REM cd ..\Excel_1_4
REM call ..\..\node_modules\.bin\api-extractor run
REM cd ..\Excel_1_3
REM call ..\..\node_modules\.bin\api-extractor run
REM cd ..\Excel_1_2
REM call ..\..\node_modules\.bin\api-extractor run
REM cd ..\Excel_1_1
REM call ..\..\node_modules\.bin\api-extractor run
REM cd ..

REM cd ..\api-extractor-inputs-word
REM call ..\node_modules\.bin\api-extractor run
REM cd ..\api-extractor-inputs-word-release\Word_1_3
REM call ..\..\node_modules\.bin\api-extractor run
REM cd ..\Word_1_2
REM call ..\..\node_modules\.bin\api-extractor run
REM cd ..\Word_1_1
REM call ..\..\node_modules\.bin\api-extractor run
REM cd ..

REM cd ..\api-extractor-inputs-onenote
REM call ..\node_modules\.bin\api-extractor run

REM cd ..\api-extractor-inputs-visio
REM call ..\node_modules\.bin\api-extractor run

REM cd ..\api-extractor-inputs-powerpoint
REM call ..\node_modules\.bin\api-extractor run

REM cd ..\api-extractor-inputs-custom-functions-runtime
REM call ..\node_modules\.bin\api-extractor run

REM cd ..\api-extractor-inputs-office-runtime
REM call ..\node_modules\.bin\api-extractor run

REM cd ..\api-extractor-inputs-outlook
REM call ..\node_modules\.bin\api-extractor run
REM cd ..\api-extractor-inputs-outlook-release\Outlook_1_7
REM call ..\..\node_modules\.bin\api-extractor run
REM cd ..\Outlook_1_6
REM call ..\..\node_modules\.bin\api-extractor run
REM cd ..\Outlook_1_5
REM call ..\..\node_modules\.bin\api-extractor run
REM cd ..\Outlook_1_4
REM call ..\..\node_modules\.bin\api-extractor run
REM cd ..\Outlook_1_3
REM call ..\..\node_modules\.bin\api-extractor run
REM cd ..\Outlook_1_2
REM call ..\..\node_modules\.bin\api-extractor run
REM cd ..\Outlook_1_1
REM call ..\..\node_modules\.bin\api-extractor run
REM cd ..
REM cd ..

pushd scripts
call node midprocessor.js
popd

call .\node_modules\.bin\api-documenter yaml --input-folder .\versioned-json\office --output-folder .\versioned-yaml\office --office

call .\node_modules\.bin\api-documenter yaml --input-folder .\versioned-json\excel --output-folder .\versioned-yaml\excel --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\versioned-json\excel_1_1 --output-folder .\versioned-yaml\excel_1_1 --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\versioned-json\excel_1_2 --output-folder .\versioned-yaml\excel_1_2 --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\versioned-json\excel_1_3 --output-folder .\versioned-yaml\excel_1_3 --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\versioned-json\excel_1_4 --output-folder .\versioned-yaml\excel_1_4 --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\versioned-json\excel_1_5 --output-folder .\versioned-yaml\excel_1_5 --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\versioned-json\excel_1_6 --output-folder .\versioned-yaml\excel_1_6 --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\versioned-json\excel_1_7 --output-folder .\versioned-yaml\excel_1_7 --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\versioned-json\excel_1_8 --output-folder .\versioned-yaml\excel_1_8 --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\versioned-json\office --output-folder .\versioned-yaml\office --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\versioned-json\outlook --output-folder .\versioned-yaml\outlook --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\versioned-json\outlook_1_1 --output-folder .\versioned-yaml\outlook_1_1 --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\versioned-json\outlook_1_2 --output-folder .\versioned-yaml\outlook_1_2 --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\versioned-json\outlook_1_3 --output-folder .\versioned-yaml\outlook_1_3 --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\versioned-json\outlook_1_4 --output-folder .\versioned-yaml\outlook_1_4 --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\versioned-json\outlook_1_5 --output-folder .\versioned-yaml\outlook_1_5 --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\versioned-json\outlook_1_6 --output-folder .\versioned-yaml\outlook_1_6 --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\versioned-json\outlook_1_7 --output-folder .\versioned-yaml\outlook_1_7 --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\versioned-json\powerpoint --output-folder .\versioned-yaml\powerpoint --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\versioned-json\onenote --output-folder .\versioned-yaml\onenote --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\versioned-json\visio --output-folder .\versioned-yaml\visio --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\versioned-json\word --output-folder .\versioned-yaml\word --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\versioned-json\word_1_1 --output-folder .\versioned-yaml\word_1_1 --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\versioned-json\word_1_2 --output-folder .\versioned-yaml\word_1_2 --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\versioned-json\word_1_3 --output-folder .\versioned-yaml\word_1_3 --office


pushd scripts
call node postprocessor.js
popd

pause