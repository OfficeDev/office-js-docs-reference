IF EXIST "node_modules" (
    rmdir "node_modules" /s /q
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

cd ..\api-extractor-inputs-word

call ..\node_modules\.bin\api-extractor run

cd ..\api-extractor-inputs-onenote

call ..\node_modules\.bin\api-extractor run

cd ..\api-extractor-inputs-visio

call ..\node_modules\.bin\api-extractor run

cd ..\api-extractor-inputs-outlook

call ..\node_modules\.bin\api-extractor run

REM cd ..\api-extractor-inputs-outlook-legacy\Outlook_1_6

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

cd ..

call .\node_modules\.bin\api-documenter yaml --input-folder .\json --office

pushd scripts
call node postprocessor.js
popd

pause