set MAJ=0
set MIN=0
set BLD=4
set REV=0

"C:\Program Files (x86)\MSBuild\14.0\Bin\msbuild.exe" buildScript.xml /t:Clean /p:Major=%MAJ% /p:Minor=%MIN% /p:Build=%BLD% /p:Revision=%REV%

"C:\Program Files (x86)\MSBuild\14.0\Bin\msbuild.exe" buildScript.xml /t:Compile /p:Major=%MAJ% /p:Minor=%MIN% /p:Build=%BLD% /p:Revision=%REV%

cd ..

dotnet pack project.json --no-build --output nupkgs

cd Build

"C:\Program Files (x86)\MSBuild\14.0\Bin\msbuild.exe" buildScript.xml /t:CleanSymbolsPkg /p:Major=%MAJ% /p:Minor=%MIN% /p:Build=%BLD% /p:Revision=%REV%

pause