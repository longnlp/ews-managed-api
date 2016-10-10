set MAJ=0
set MIN=0
set BLD=1
set REV=0

"C:\Program Files (x86)\MSBuild\14.0\Bin\msbuild.exe" buildScript.xml /t:Clean /p:Major=%MAJ% /p:Minor=%MIN% /p:Build=%BLD% /p:Revision=%REV%

pause