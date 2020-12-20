Add-Type -Path "$PSScriptRoot\bin\Microsoft.WindowsAPICodePack.Shell.dll"

foreach ($f in Get-ChildItem -Path "$PSScriptRoot\functions" -Filter *.ps1) {
    . $f.FullName
}

#https://stackoverflow.com/questions/5337683/how-to-set-extended-file-properties
#https://www.nuget.org/packages/Microsoft-WindowsAPICodePack-Shell/