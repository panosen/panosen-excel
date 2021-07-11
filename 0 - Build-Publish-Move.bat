@echo off

dotnet restore

dotnet build --no-restore -c Release

move /Y Panosen.Excel\bin\Release\Panosen.Excel.*.nupkg D:\LocalSavoryNuget\

pause