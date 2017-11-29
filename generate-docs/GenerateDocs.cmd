call npm install

del package-lock.json

cd excel

call ..\node_modules\.bin\api-extractor run

cd ..


call .\node_modules\.bin\api-documenter yaml --input-folder .\json --office

rem call D:\GitRepos\wbt2\libraries\api-documenter\ad.cmd yaml --input-folder .\json --office

pause
