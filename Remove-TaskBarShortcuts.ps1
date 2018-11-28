[CmdletBinding(SupportsShouldProcess=$True)]
param()

$AllUsersPath = $env:ALLUSERSPROFILE

$scnames = @("Excel (desktop)","Excel 2013", "Excel 2016",
"Lync (desktop)", "Lync 2013","Outlook (desktop)","Outlook 2013", 
"Outlook 2016", "PowerPoint (desktop)","PowerPoint 2013", "PowerPoint 2016", 
"Skype for Business 2015", "Word (desktop)","Word 2013", "Word 2016",
"Access","Excel","Outlook","PowerPoint","Word","Skype for Business")

$sclinks = $scnames | %{"$_.lnk"}

if (Test-Path "c:\users") {
    $UsersPath = "c:\users"
}
else {
    $UsersPath = "c:\documents and settings"
}

$profiles = Get-ChildItem -Path $UsersPath
Write-Verbose "$($profiles.Count) profiles found"

foreach ($profile in $profiles) {
    $userName = $profile.Name
    Write-Verbose "profile: $userName"
    $BasePath = Join-Path -Path $UsersPath -ChildPath $profile.Name
    $LinkPath = Join-Path -Path $BasePath -ChildPath "AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"
    if (Test-Path $LinkPath) {
        #Write-Verbose $LinkPath
        $shortcuts = Get-ChildItem -Path $LinkPath -Filter "*.lnk"
        #Write-Verbose "$($shortcuts.Count) shortcuts were found: $($profile.Name)"
        foreach ($sc in $shortcuts) {
            if ($sclinks -contains $sc.Name) {
                Write-Verbose "--Found: $($sc.Name)"
                $scPath = Join-Path -Path $LinkPath -ChildPath $sc.Name
                try {
                    Remove-Item -Path $scPath -Force
                    Write-Verbose "--Deleted"
                }
                catch {
                    Write-Error $Error[0].Exception.Message
                }
            }
        }
    }
}
