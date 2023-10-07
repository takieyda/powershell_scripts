$url = 'https://grc.sc'
$shortcuts = [ordered]@{}
$range = 1..999

Function makeRequest {
    $cnt = 0
    ForEach ($sc in $range) {
        $cnt += 1
        Write-Progress -Activity "Requesting" -Status "$cnt/$(($range).Count) complete" -PercentComplete ($cnt/($range).Count * 100)
        $response = Invoke-WebRequest -Uri $url/$sc -MaximumRedirection 0 -ErrorAction SilentlyContinue
        If ($response.StatusCode -eq 301) {
            $title = (Invoke-WebRequest -Uri $url/$sc -UseBasicParsing).ParsedHtml.Title
            If ($title -eq $null) {
                $title = (($response).Headers.Location -split '/')[-1]
            }

            Write-Host "$cnt :: $sc :: $($response.StatusCode) :: $title"
            cacheShortcuts $sc $url/$sc $response.Headers.Location $title
        } ElseIf (($response.StatusCode -eq 302) -and ($response.Headers.Location -eq 'https://grc.sc')) {
            $pass #Write-Host 'None' -ForegroundColor green
        } ElseIf ($response.StatusCode -eq (400..499)) {
            Write-Host "Error: $response.StatusCode :: $url/$sc" -ForegroundColor Red
        }
    }
}

Function cacheShortcuts($sc,$shortcut,$destination,$title) {
    $shortcuts["$sc"] = [ordered]@{ 
        "sc"          = $sc;
        "shortcut"    = $shortcut;
        "destination" = $destination;
        "title"       = $title
    }
}

Function saveFile {
    Write-Host "`n`n========================`n`nSaving shortcuts."
    If ($shortcuts -ne $null ) {
        ForEach ($item in $shortcuts.Keys) {
            New-Object -TypeName psobject -Property $shortcuts.$item `
                | Select-Object sc,shortcut,destination,title `
                | Export-CSV -Path ".\grc_shortcuts.csv" -NoTypeInformation -Append
        }
    } Else {
        Write-host "Nothing to save." -ForegroundColor Red
    }
}              

Try { makeRequest }
Catch { Write-Host $_ }
Finally { saveFile }
