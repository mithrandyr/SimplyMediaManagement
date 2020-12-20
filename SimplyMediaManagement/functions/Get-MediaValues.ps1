function Get-MediaValues {
    [cmdletBinding(DefaultParameterSetName="object")]
    param([parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName, ParameterSetName="object")]
            [System.IO.FileInfo]$file
        , [parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName, ParameterSetName="path", Position = 0)]
            [Alias("FullName")]
            [string]$path)

    process {
        if($PSCmdlet.ParameterSetName -eq "object") { $path = $file.FullName }
        
        if(-not (Test-Path $path -PathType Leaf)) { Write-Warning -Message "Invalid Path -> $path" }
        else {
            if($PSCmdlet.ParameterSetName -eq "path") { $file = Get-Item -Path $path }
            $d = [Microsoft.WindowsAPICodePack.Shell.ShellFile]::FromFilePath($file.FullName)
            
            [PSCustomObject]@{
                FilePath = $file.FullName
                Series = $d.Properties.System.Title.Value
                Title = $d.Properties.System.Media.SubTitle.Value
                Genres = ($d.Properties.System.Music.Genre.Value) -join ";"
                Year = $d.Properties.System.Media.Year.Value
                Artists = ($d.Properties.System.Author.Value) -join ";"
                Publisher = $d.Properties.System.Media.Publisher.Value
                Rating = {param($_) if($_.Value -ne $null) {($_.Value / 25) + 1} }.Invoke($d.Properties.System.Rating)[0]
            }
        }
    }
}

Export-ModuleMember -Function Get-MediaValues