call npm install

del package-lock.json

cd api-extractor-inputs

call ..\node_modules\.bin\api-extractor run

cd ..

call .\node_modules\.bin\api-documenter yaml --input-folder .\api-documenter-inputs --office

rem call D:\GitRepos\wbt2\libraries\api-documenter\ad.cmd yaml --input-folder .\json --office

pause
