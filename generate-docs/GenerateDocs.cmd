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
call node preprocessor.js
popd

del package-lock.json

cd api-extractor-inputs-office

call ..\node_modules\.bin\api-extractor run

cd ..\api-extractor-inputs-excel

call ..\node_modules\.bin\api-extractor run

cd ..\api-extractor-inputs-excel-release\Excel_1_8

call ..\..\node_modules\.bin\api-extractor run

cd ..\Excel_1_7

call ..\..\node_modules\.bin\api-extractor run

cd ..\Excel_1_6

call ..\..\node_modules\.bin\api-extractor run

cd ..\Excel_1_5

call ..\..\node_modules\.bin\api-extractor run

cd ..\Excel_1_4

call ..\..\node_modules\.bin\api-extractor run

cd ..\Excel_1_3

call ..\..\node_modules\.bin\api-extractor run

cd ..\Excel_1_2

call ..\..\node_modules\.bin\api-extractor run

cd ..\Excel_1_1

call ..\..\node_modules\.bin\api-extractor run

cd ..

cd ..\api-extractor-inputs-word

call ..\node_modules\.bin\api-extractor run

cd ..\api-extractor-inputs-word-release\Word_1_3

call ..\..\node_modules\.bin\api-extractor run

cd ..\Word_1_2

call ..\..\node_modules\.bin\api-extractor run

cd ..\Word_1_1

call ..\..\node_modules\.bin\api-extractor run

cd ..

cd ..\api-extractor-inputs-onenote

call ..\node_modules\.bin\api-extractor run

cd ..\api-extractor-inputs-visio

call ..\node_modules\.bin\api-extractor run

cd ..\api-extractor-inputs-powerpoint

call ..\node_modules\.bin\api-extractor run

cd ..\api-extractor-inputs-outlook

call ..\node_modules\.bin\api-extractor run

cd ..\api-extractor-inputs-custom-functions-runtime

call ..\node_modules\.bin\api-extractor run

cd ..\api-extractor-inputs-office-runtime

call ..\node_modules\.bin\api-extractor run

cd ..\api-extractor-inputs-outlook-release\Outlook_1_7

call ..\..\node_modules\.bin\api-extractor run

cd ..\Outlook_1_6

call ..\..\node_modules\.bin\api-extractor run

cd ..\Outlook_1_5

call ..\..\node_modules\.bin\api-extractor run

cd ..\Outlook_1_4

call ..\..\node_modules\.bin\api-extractor run

cd ..\Outlook_1_3

call ..\..\node_modules\.bin\api-extractor run

cd ..\Outlook_1_2

call ..\..\node_modules\.bin\api-extractor run

cd ..\Outlook_1_1

call ..\..\node_modules\.bin\api-extractor run

cd ..

cd ..

pushd scripts
call node midprocessor.js
popd

call .\node_modules\.bin\api-documenter yaml --input-folder .\json --office

call .\node_modules\.bin\api-documenter yaml --input-folder .\versioned-json\excel --output-folder .\versioned-yaml\excel --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\versioned-json\excel_1_1 --output-folder .\versioned-yaml\excel_1_1 --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\versioned-json\excel_1_2 --output-folder .\versioned-yaml\excel_1_2 --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\versioned-json\excel_1_3 --output-folder .\versioned-yaml\excel_1_3 --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\versioned-json\excel_1_4 --output-folder .\versioned-yaml\excel_1_4 --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\versioned-json\excel_1_5 --output-folder .\versioned-yaml\excel_1_5 --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\versioned-json\excel_1_6 --output-folder .\versioned-yaml\excel_1_6 --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\versioned-json\excel_1_7 --output-folder .\versioned-yaml\excel_1_7 --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\versioned-json\excel_1_8 --output-folder .\versioned-yaml\excel_1_8 --office
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

pushd scripts
call node postprocessor.js
popd

pause