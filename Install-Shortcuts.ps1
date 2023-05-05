# install desktop shortcuts
# based on: https://www.nickydewestelinck.be/2021/03/09/how-to-deploy-custom-url-shortcuts-with-microsoft-intune/

[CmdletBinding()]
param (
    # path to JSON file with shortcut variables
    [Parameter()]
    [string]
    $shortcutFile = '.\shortcuts.json',

    [Parameter()]
    [switch]
    $Install,

    [Parameter()]
    [switch]
    $Uninstall
)
    

begin {
    enum ShortcutType {
        URL
        Link
    }
    
    class IntuneShortcut {
        # shortcut type
        [ShortcutType]
        $Type
    
        # source of the icon. can be a URL or file
        [string]
        $SourceIcon = $null
    
        # local name of the icon
        [string]
        $IconName = $null
    
        # the target path for the shortcut
        [string]
        $TargetPath = $null
    
        # shortcut name
        [string]
        $ShortcutName = $null
    
        # where do shortcuts go
        [string] hidden
        $ShortcutPath = "$([System.Environment]::GetFolderPath("Desktop"))"
    
        # where do the icon images go
        [string] hidden
        $IconPath = "$([System.Environment]::GetFolderPath("CommonPictures"))\icons"
    
        # shortcut with custom icon
        IntuneShortcut(
            $Type,
            [string]$SourceIcon,
            [string]$IconName,
            [string]$TargetPath,
            [string]$ShortcutName
        ) {
            $this.Type = $type
    
            # ignore these for URL type, since they are not supported
            if ($type -eq [ShortcutType]"Link") {
                $this.SourceIcon   = $SourceIcon
                $this.IconName     = $IconName
            }
            
            $this.TargetPath   = $TargetPath
            $this.ShortcutName = $ShortcutName
        }
    
        # shortcut without icon
        IntuneShortcut(
            $Type,
            [string]$TargetPath,
            [string]$ShortcutName
        ) {
            $this.Type = $type
            $this.TargetPath   = $TargetPath
            $this.ShortcutName = $ShortcutName
        }
    
        # used for removing shortcut
        IntuneShortcut(
            [string]$ShortcutName
        ) {
            $this.ShortcutName = $ShortcutName
        }
    
        ### GETTERS ###
        [string]
        GetShortcutPath() {
            return $this.ShortcutPath
        }
    
        [string]
        GetIconPath() {
            return $this.IconPath
        }
    
    
        ### SETTERS ###
        SetShortcutPath([string]$SP) {
            if ( -NOT (Test-Path "$SP" -IsValid) ) {
                Write-Error "The shortcut path is invalid: $($Error[0].ToString())"
            } else {
                $this.ShortcutPath = $SP
            }        
        }
    
        SetIconPath([string]$IconPath) {
            if ( -NOT (Test-Path "$IconPath" -IsValid) ) {
                Write-Error "The shortcut path is invalid: $($Error[0].ToString())"
            } else {
                $this.IconPath = $IconPath
            }
        }
    
    
        ### WORK ###
        [string] hidden 
        DownloadIcon([string]$URL) {
            Write-Verbose "[IntuneShortcut].DownloadIcon - Begin"
            # force TLS 1.2 or 1.3
            [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12, [System.Net.SecurityProtocolType]::Tls13
    
            # download the icon
            try {
                Write-Verbose "[IntuneShortcut].DownloadIcon - Downloading icon."
                $null = Invoke-WebRequest -Uri $URL -OutFile "$($this.IconPath)\$($this.IconName)"
                Write-Verbose "[IntuneShortcut].DownloadIcon - Icon saved to $($this.IconPath)\$($this.IconName)"
                Write-Verbose "[IntuneShortcut].DownloadIcon - End"
                return "$($this.IconPath)\$($this.IconName)"
            } catch {
                Write-Error "Failed to download the icon: $_"
                Write-Verbose "[IntuneShortcut].DownloadIcon - End"
                return $null   
            }
        }
    
        AddShortcut() {
            Write-Verbose "[IntuneShortcut].AddShortcut - Begin"
            # the shell object is used to create the shortcut
            Write-Verbose "[IntuneShortcut].AddShortcut - Create shell object."
            $new_object = New-Object -ComObject WScript.Shell
    
            # create the shortcut
            Write-Verbose "[IntuneShortcut].AddShortcut - shortcutpath: $($this.ShortcutPath)\$($this.ShortcutName)"
            $source = $new_object.CreateShortcut("$($this.ShortcutPath)\$($this.ShortcutName)")
    
            # add the target
            $source.TargetPath = $this.TargetPath
            Write-Verbose "[IntuneShortcut].AddShortcut - Added target: $($this.TargetPath)"
           
            # URLs types do not support custom icons, so only do custom stuff if this is a Link type shortcut
            if ($this.Type -eq [ShortcutType]"Link") {
                Write-Verbose "[IntuneShortcut].AddShortcut - Processing a link type shortcut."
                $isLclFile = Get-Item "$($this.SourceIcon)" -EA SilentlyContinue

                if ( -NOT $isLclFile ) {
                    if (($this.SourceIcon)[0] -eq '.') {
                        Write-Verbose "[IntuneShortcut].AddShortcut - Trying PSScriptRoot"
                        $tmp = "$PSScriptRoot\$($this.SourceIcon.TrimStart('.'))"

                        $isLclFile = Get-Item "$tmp" -EA SilentlyContinue
                        if ($isLclFile) {
                            Write-Verbose "[IntuneShortcut].AddShortcut - PSScriptRoot result: $($isLclFile.FullName)"
                        } else {
                            Write-Verbose "[IntuneShortcut].AddShortcut - PSScriptRoot failed."
                        }
                    }
                }

                if ($isLclFile) {
                    Write-Verbose "[IntuneShortcut].AddShortcut - Using SourceIcon as icon: $($this.SourceIcon)"
                    $icon = $isLclFile.FullName
                } else {
                    Write-Verbose "[IntuneShortcut].AddShortcut - Attempting to downloading the icon."
                    $icon = $this.DownloadIcon($this.SourceIcon)
                }

                # create the icon storage path if it does not exist
                if ( -NOT (Test-Path "$($this.IconPath)") ) {
                    Write-Verbose "[IntuneShortcut].AddShortcut - Creating icon directory."
                    New-Item "$($this.IconPath)" -Type Directory
                }

                # save the icon to the shortcut if the ico file was successfully created
                Write-Verbose "[IntuneShortcut].AddShortcut - Icon path: $icon"
                if ( -NOT [string]::IsNullOrEmpty($icon) ) {
                    Write-Verbose "[IntuneShortcut].AddShortcut - Adding icon ($icon) to shortcut at path $($this.GetIconPath())."
                    $null = Copy-Item "$icon" "$($this.GetIconPath())" -Force
                    $source.IconLocation = $icon
                }
            }
    
            # save/commit the shortcut
            Write-Verbose "[IntuneShortcut].AddShortcut - Creating icon."
            $source.Save()
            Write-Verbose "[IntuneShortcut].AddShortcut - End"
        }
    
        RemoveShortcut() {
            $scPath = "$($this.ShortcutPath)\$($this.ShortcutName)"
    
            if ((Test-Path "$scPath")) {
                $null = Remove-Item $scPath -Force
            } else {
                Write-Warning "Shortcut not found at $scPath."
            }
        }
    }
    
    # read and convert
    $scObj = Get-Content $shortcutFile | ConvertFrom-Json
    Write-Verbose "$($scObj | Format-List * | Out-String)"
}


process {
    if ($Install.IsPresent) {
        foreach ($obj in $scObj) {
            if ($obj.Type -eq "URL") {
                $tmp = [IntuneShortcut]::new($obj.Type, $obj.Target, $obj.Name)
                $tmp.AddShortcut()
                Clear-Variable tmp -EA SilentlyContinue
            } elseif ($obj.Type -eq "Link") {
                $tmp = [IntuneShortcut]::new($obj.Type, $obj.SourceIcon, $obj.IconName, $obj.Target, $obj.Name)
                $tmp.AddShortcut()
                Clear-Variable tmp -EA SilentlyContinue
            }
        }
    } elseif ($Uninstall.IsPresent) {
        foreach ($obj in $scObj) {
            $tmp = [IntuneShortcut]::new($obj.Name)
            $tmp.RemoveShortcut()
            Clear-Variable tmp -EA SilentlyContinue
        }
    }
}

end{}