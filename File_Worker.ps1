# DEFINE PARAMETERS
param (
    [string]$fPath,
    [string]$act
)

# DEFINE KEY FUNCTIONS

function Get-ValidInput { #An absolute banger of a function, I love it like my own child
    param (
        [string]$prompt,
        [array]$validValues
    )
    do {
        $entry = Read-Host -Prompt $prompt
        if ($entry -in $validValues) {
            return $entry
        } else {
            Write-Host "Invalid input."
        }
    } while ($true)
}

function Get-FolderPath {
    param (
        [string]$prompt
    )
    do {
        $path = Read-Host -Prompt $prompt
        if (Test-Path -Path $path -PathType Container) {
            return $path
        } else {
            Show-NonexistentFolderMessage
            return $null
        }
    } while ($true)
}

function Show-NonexistentFolderMessage {
    Write-Host "The folder does not exist."
}

function Select-Action {
    $act = Get-ValidInput -prompt "Choose an action: Organize (1), Convert (2), List files (3), Change Folder (4), Exit (5)" -validValues @("1", "2", "3", "4", "5")
    return $act
}

function Invoke-EnsureFolderAndMoveFile {
    param (
        [string]$targetFolder,
        [string]$filePath      
    )
    if (-not (Test-Path -Path $targetFolder)) {
        New-Item -ItemType Directory -Path $targetFolder | Out-Null
    }
    Move-Item -Path $filePath -Destination $targetFolder
}

function Invoke-PromptAndMoveToOther {
    param (
        [string]$filePath,
        [string]$folderPath
    )

    $otherFolder = Join-Path -Path $folderPath -ChildPath "Other"
    if (-not (Test-Path -Path $otherFolder)) {
        New-Item -ItemType Directory -Path $otherFolder | Out-Null
    }
    Move-Item -Path $filePath -Destination $otherFolder
}

# DEFINE OPERATING FUNCTIONS

function Show-DirectoryContents {
    param (
        [string]$path
    )
    Get-ChildItem -Path $fPath | ForEach-Object {
        $size = $_.Length
        $formattedSize = if ($size -lt 1KB) {
            "$size B"
        } elseif ($size -lt 1MB) {
            "$([math]::Round($size / 1KB, 2)) KB"
        } elseif ($size -lt 1GB) {
            "$([math]::Round($size / 1MB, 2)) MB"
        } else {
            "$([math]::Round($size / 1GB, 2)) GB"
        }
        Write-Host "Name: $($_.Name), Size: $formattedSize, Last Modified: $($_.LastWriteTime)"
    }
}
function Invoke-OrganizeFiles {
    param (
        [string]$path
    )
    # The true hardworker here
    function Invoke-OrganizeInFolder {
        param (
            [string]$folderPath
        )

        switch ($grandOrgType) {
            "1" { # Organizing by date
                $orgType = Get-ValidInput -prompt "Creation Date (1), Last Modified Date (2), Last Accessed Date (3)" -validValues @("1", "2", "3")
                $property = switch ($orgType) {
                    "1" { "CreationTime" }
                    "2" { "LastWriteTime" }
                    "3" { "LastAccessTime" }
                }

                Get-ChildItem -Path $folderPath -File | ForEach-Object {
                    $dateFolder = $_.$property.ToString("yyyy-MM-dd")
                    $targetFolder = Join-Path -Path $folderPath -ChildPath $dateFolder
                    if ($targetFolder -notin $existingFolders) {
                        Invoke-EnsureFolderAndMoveFile -targetFolder $targetFolder -filePath $_.FullName
                    }
                }
            }
            "2" { # Organizing by type
                Get-ChildItem -Path $folderPath -File | ForEach-Object {
                    $typeFolder = $_.Extension.TrimStart(".")
                    if (-not $typeFolder) { $typeFolder = "NoExtension" }
                    $targetFolder = Join-Path -Path $folderPath -ChildPath $typeFolder
                    if ($targetFolder -notin $existingFolders) {
                        Invoke-EnsureFolderAndMoveFile -targetFolder $targetFolder -filePath $_.FullName
                    }
                }
            }
            "3" { # Organizing by size
                Get-ChildItem -Path $folderPath -File | ForEach-Object {
                    $size = $_.Length
                    $sizeFolder = if ($size -lt 1MB) {
                        "Small (-1MB)"
                    } elseif ($size -lt 10MB) {
                        "Medium (1MB-10MB)"
                    } elseif ($size -lt 100MB) {
                        "Large (10MB-100MB)"
                    } else {
                        "Huge (100MB+)"
                    }
                    $targetFolder = Join-Path -Path $folderPath -ChildPath $sizeFolder
                    if ($targetFolder -notin $existingFolders) {
                        Invoke-EnsureFolderAndMoveFile -targetFolder $targetFolder -filePath $_.FullName
                    }
                }
            }
            "4" { # Organizing by name
                Get-ChildItem -Path $folderPath -File | ForEach-Object {
                    $firstCharacter = $_.Name.Substring(0, 1) # Get the first character of the file name
                    $sanitizedCharacter = $firstCharacter -replace '[\\/:*?"<>|]', "_" # Replace invalid characters with an underscore
                    $targetFolder = Join-Path -Path $folderPath -ChildPath $sanitizedCharacter
                    
                    if ($targetFolder -notin $existingFolders) {
                        Invoke-EnsureFolderAndMoveFile -targetFolder $targetFolder -filePath $_.FullName
                    }
                }
            }
            "5" { # Multimedia advanced
                switch ($mediaType) {
                    "1" { # Organizing by resolution (Images)
                        # Check if ImageMagick is installed
                        if (-not (Get-Command "magick" -ErrorAction SilentlyContinue)) {
                            Write-Host "ImageMagick is required for this operation. Please install it and try again."
                            return
                        }

                        Get-ChildItem -Path $folderPath -File | ForEach-Object {
                            if ($_.Extension -match "\.(jpg|jpeg|png|gif)$") {
                                # Use ImageMagick to get image resolution
                                $imageInfo = & magick identify -format "%wx%h" $_.FullName
                                if ($imageInfo -match "(\d+)x(\d+)") {
                                    $width = [int]$matches[1]
                                    $height = [int]$matches[2]

                                    # Determine resolution category
                                    $resolutionFolder = if ($width -le 640 -and $height -le 480) {
                                        "LowRes (640x480-)"
                                    } elseif ($width -le 1280 -and $height -le 720) {
                                        "HD (1280x720-)"
                                    } elseif ($width -le 1920 -and $height -le 1080) {
                                        "FullHD (1920x1080-)"
                                    } elseif ($width -le 2560 -and $height -le 1440) {
                                        "2K (2560x1440-)"
                                    } elseif ($width -le 3840 -and $height -le 2160) {
                                        "4K (3840x2160-)"
                                    } elseif ($width -le 7680 -and $height -le 4320) {
                                        "8K (7680x4320-)"
                                    } else {
                                        "HigherRes (8K+)"
                                    }

                                    # Move file to the appropriate folder
                                    $targetFolder = Join-Path -Path $folderPath -ChildPath $resolutionFolder
                                    if ($targetFolder -notin $existingFolders) {
                                        Invoke-EnsureFolderAndMoveFile -targetFolder $targetFolder -filePath $_.FullName
                                    }
                                }
                            } else {
                                if ($moveToOther -eq "Y") {
                                    Invoke-PromptAndMoveToOther -filePath $_.FullName -folderPath $folderPath
                                }
                            }
                        }
                    }
                    "2" { # Organizing by duration (Audio/Video)
                        # Check if ffmpeg is installed
                        if (-not (Get-Command "ffmpeg" -ErrorAction SilentlyContinue)) {
                            Write-Host "FFmpeg is required for this operation. Please install it and try again."
                            return
                        }

                        Get-ChildItem -Path $folderPath -File | ForEach-Object {
                            if ($_.Extension -match "\.(mp3|mp4|wav|avi)$") {
                                # Use ffmpeg to get duration
                                $durationInfo = & ffmpeg -i $_.FullName 2>&1 | Select-String "Duration"
                                if ($durationInfo -match "Duration: (\d+):(\d+):(\d+)") {
                                    $hours = [int]$matches[1]
                                    $minutes = [int]$matches[2]
                                    $seconds = [int]$matches[3]
                                    $totalSeconds = ($hours * 3600) + ($minutes * 60) + $seconds

                                    # Determine duration category
                                    $durationFolder = if ($totalSeconds -le 300) {
                                        "Short (0-5 min)"
                                    } elseif ($totalSeconds -le 1800) {
                                        "Medium (5-30 min)"
                                    } elseif ($totalSeconds -le 3600) {
                                        "Long (30-60 min)"
                                    } else {
                                        "Very Long (60+ min)"
                                    }

                                    # Move file to the appropriate folder
                                    $targetFolder = Join-Path -Path $folderPath -ChildPath $durationFolder
                                    if ($targetFolder -notin $existingFolders) {
                                        Invoke-EnsureFolderAndMoveFile -targetFolder $targetFolder -filePath $_.FullName
                                    }
                                }
                            } else {
                                if ($moveToOther -eq "Y") {
                                    Invoke-PromptAndMoveToOther -filePath $_.FullName -folderPath $folderPath
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    # Ask for organization type
    $grandOrgType = Get-ValidInput -prompt "By Date (1), By Type (2), By Size (3), By Name (4), Multimedia advanced (5)" -validValues @("1", "2", "3", "4", "5")

    # If Multimedia advanced is selected, ask for media type
    $mediaType = $null
    if ($grandOrgType -eq "5") {
        $mediaType = Get-ValidInput -prompt "Resolution (Images) (1), Duration (Audio/Video) (2)" -validValues @("1", "2")
        $moveToOther = Get-ValidInput -prompt "Invalid files are possible. Move to Other? Yes (Y), No (N)" -validValues @("Y", "N")
    }

    # Check if there are any subfolders
    $subfolders = Get-ChildItem -Path $path -Directory -ErrorAction SilentlyContinue
    $processSubfolders = $false
    if ($subfolders) {
        # Ask for recursion option only if subfolders exist
        $includeSubfolders = Get-ValidInput -prompt "Include subfolders? Yes (Y), No (N)" -validValues @("Y", "N")
        $processSubfolders = $includeSubfolders -eq "Y"
    }

    # Get original subfolders if recursion is enabled
    if ($processSubfolders) {
        $originalSubfolders = Get-ChildItem -Path $path -Directory -Recurse
    }

    # Process the main folder
    Invoke-OrganizeInFolder -folderPath $path -grandOrgType $grandOrgType -mediaType $mediaType

    # Process subfolders if recursion is enabled
    if ($processSubfolders) {
        $originalSubfolders | ForEach-Object {
            Invoke-OrganizeInFolder -folderPath $_.FullName -grandOrgType $grandOrgType -mediaType $mediaType
        }
    }

    Write-Host "Files in $path have been organized."
}

function Invoke-ConvertFiles {
    param (
        [string]$path
    )
}

# MAIN SCRIPT

# Check the provided folder path
if ($fPath) {
    if (-not (Test-Path -Path $fPath -PathType Container)) {
        Show-NonexistentFolderMessage
        exit
    }
} else {
    $fPath = Get-FolderPath -prompt "Path to the working directory"
    if (-not $fPath) {
        exit
    }
}

# Ask for action, if not provided or invalid
if (-not $act -or $act -notin @("1", "2", "3", "4", "5")) {
    $act = Select-Action
}

# Perform the action
do {
    switch ($act) {
        "1" {
            Invoke-OrganizeFiles -path $fPath
            $act = Select-Action
        }
        "2" {
            Invoke-ConvertFiles -path $fPath
            $act = Select-Action
        }
        "3" {
            Show-DirectoryContents -path $fPath
            $act = Select-Action
        }
        "4" {
            $NewfPath = Get-FolderPath -prompt "Path to the new working directory"
            if ($NewfPath) {
                $fPath = $NewfPath
            }
            $act = Select-Action
        }
        "5" {
            exit
        }
    }
} while ($true)
