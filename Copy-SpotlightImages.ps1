<#
    .SYNOPSIS
        Copy Microsoft spotlight wallpapers to specified folder.
    .DESCRIPTION
        Script will copy Microsoft spotlight pictures to specified folder.
        After that You can set your desktop settings to read from the folder.
        This powershell script can be then scheduled to run daily via a Task Scheduler.
        Script will automatically remove spotlight images older than 7 days.
    .EXAMPLE
        PS C:\> .\Copy-SpotlightImages.ps1 -Path "$env:UserProfile\OneDrive\Pictures\wallpaper\Spotlight"
        Copy Spotlight images to a folder on OneDrive.
    .NOTES
        Note that by default script copies the files to "$env:UserProfile\OneDrive\Pictures\Wallpaper\Spotlight"
#>

#region Init
[CmdletBinding()]
param (
    [ValidateScript( { Test-Path $_ } )]
    [string] $Path = "$env:UserProfile\OneDrive\Pictures\Wallpaper\Spotlight"
)

$SpotlightPath = "$env:LocalAppData\Packages\Microsoft.Windows.ContentDeliveryManager_cw5n1h2txyewy\LocalState\Assets"
#endregion Init

#region Copy
try {
    Write-Output 'Take the big images only'
    $SpotlightFiles = Get-ChildItem -Path $SpotlightPath | Where-Object { $_.Length -gt 1kb }

    if ($SpotlightFiles) {
        Write-Output 'Copy new files'
        $SpotlightFiles | ForEach-Object { Copy-Item -Path $_.FullName -Destination $Path\$($_.Name).jpg }

        Write-Output "Clean up everything that doesn't have required size or too old"
        Get-Item $Path\*.jpg | ForEach-Object {
            $Namespace = (New-Object -ComObject Shell.Application).Namespace($Path)
            $Item = $Namespace.ParseName($_.Name)
            $Size = $Namespace.GetDetailsOf($Item, 31)

            if ($Size -match '(\d+) x (\d+)') {
                $Width = [int]($Matches[1])
                $Height = [int]($Matches[2])
            }

            if (!$Size -or
                ($Width -lt 1920 -and $Height -lt 1080) -or
                ($_.LastWriteTime -le (Get-Date).AddDays(-7))
            ) {
                Write-Output "Removing $($_.Name)"
                Remove-Item $_ -Force
            }
        }
    }
} catch {
    throw "Error copying the files from spotlight folder: $($Error[0])"
}
#endregion Copy