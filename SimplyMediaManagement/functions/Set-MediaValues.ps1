function Set-MediaValues {
    [cmdletBinding()]
    param([parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)][string]$FilePath
        , [parameter(ValueFromPipelineByPropertyName)][string]$Title
        , [parameter(ValueFromPipelineByPropertyName)][string]$Series
        , [parameter(ValueFromPipelineByPropertyName)][string]$Genres
        , [parameter(ValueFromPipelineByPropertyName)][string]$Year
        , [parameter(ValueFromPipelineByPropertyName)][string]$Artists
        , [parameter(ValueFromPipelineByPropertyName)][string]$Publisher
        , [parameter(ValueFromPipelineByPropertyName)][ValidateSet(0,1,2,3,4,5)][int]$Rating
    )

    process {
        if(-not (Test-Path $FilePath -PathType Leaf)) { Write-Warning -Message "Invalid Path -> $FilePath" }
        else {
            $file = Get-Item -Path $FilePath
            $d = [Microsoft.WindowsAPICodePack.Shell.ShellFile]::FromFilePath($file.FullName)
            
            if($PSBoundParameters.Keys.Contains("Title")) { $d.Properties.System.Media.Subtitle.Value = $Title }
            if($PSBoundParameters.Keys.Contains("Series")) { $d.Properties.System.Title.Value = $Series }
            if($PSBoundParameters.Keys.Contains("Genres")) { $d.Properties.System.Music.Genre.Value = $Genres }
            if($PSBoundParameters.Keys.Contains("Year")) { $d.Properties.System.Media.Year.Value = $Year }
            if($PSBoundParameters.Keys.Contains("Artists")) { $d.Properties.System.Author.Value = $Artists }
            if($PSBoundParameters.Keys.Contains("Publisher")) { $d.Properties.System.Media.Publisher.Value = $Publisher }
            if($PSBoundParameters.Keys.Contains("Rating")) {
                if($Rating -eq 0) { $d.Properties.System.Rating.ClearValue() }
                else{ $d.Properties.System.Rating.Value = ($Rating - 1) * 25 }
            }
        }
    }
}

Export-ModuleMember -Function Set-MediaValues