dotnet restore

dotnet build --no-restore -c Debug

copy-pdb-to-nuget debug

pause