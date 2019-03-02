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
cd ..\api-extractor-inputs-excel-release\excel_1_8
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
cd ..\api-extractor-inputs-outlook-release\outlook_1_7
call ..\..\node_modules\.bin\api-extractor run
cd ..\outlook_1_6
call ..\..\node_modules\.bin\api-extractor run
cd ..\outlook_1_5
call ..\..\node_modules\.bin\api-extractor run
cd ..\outlook_1_4
call ..\..\node_modules\.bin\api-extractor run
cd ..\outlook_1_3
call ..\..\node_modules\.bin\api-extractor run
cd ..\outlook_1_2
call ..\..\node_modules\.bin\api-extractor run
cd ..\outlook_1_1
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
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\excel_1_1 --output-folder .\yaml\excel_1_1 --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\excel_1_2 --output-folder .\yaml\excel_1_2 --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\excel_1_3 --output-folder .\yaml\excel_1_3 --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\excel_1_4 --output-folder .\yaml\excel_1_4 --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\excel_1_5 --output-folder .\yaml\excel_1_5 --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\excel_1_6 --output-folder .\yaml\excel_1_6 --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\excel_1_7 --output-folder .\yaml\excel_1_7 --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\excel_1_8 --output-folder .\yaml\excel_1_8 --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\office --output-folder .\yaml\office --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\outlook --output-folder .\yaml\outlook --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\outlook_1_1 --output-folder .\yaml\outlook_1_1 --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\outlook_1_2 --output-folder .\yaml\outlook_1_2 --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\outlook_1_3 --output-folder .\yaml\outlook_1_3 --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\outlook_1_4 --output-folder .\yaml\outlook_1_4 --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\outlook_1_5 --output-folder .\yaml\outlook_1_5 --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\outlook_1_6 --output-folder .\yaml\outlook_1_6 --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\outlook_1_7 --output-folder .\yaml\outlook_1_7 --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\powerpoint --output-folder .\yaml\powerpoint --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\onenote --output-folder .\yaml\onenote --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\visio --output-folder .\yaml\visio --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\word --output-folder .\yaml\word --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\word_1_1 --output-folder .\yaml\word_1_1 --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\word_1_2 --output-folder .\yaml\word_1_2 --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\word_1_3 --output-folder .\yaml\word_1_3 --office


pushd scripts
call node postprocessor.js
popd

pause