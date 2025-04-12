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
        [string]$folderPath,
        [string]$actionType
    )

    if ($actionType -eq "Originals") {
        $targetFolder = Join-Path -Path $folderPath -ChildPath "Originals"
    } elseif ($actionType -eq "Failed") {
        $targetFolder = Join-Path -Path $folderPath -ChildPath "Failed"
    } else {
        $targetFolder = Join-Path -Path $folderPath -ChildPath "Other"
    }
    
    if (-not (Test-Path -Path $targetFolder)) {
        New-Item -ItemType Directory -Path $targetFolder | Out-Null
    }
    Move-Item -Path $filePath -Destination $targetFolder
}

function Get-Subfolders {
    param (
        [string]$path
    )

    $subfolders = Get-ChildItem -Path $path -Directory -ErrorAction SilentlyContinue
    $processSubfolders = $false

    if ($subfolders) {
        $includeSubfolders = Get-ValidInput -prompt "Include subfolders? Yes (Y), No (N)" -validValues @("Y", "N")
        $processSubfolders = $includeSubfolders -eq "Y"
    }

    if ($processSubfolders) {
        return Get-ChildItem -Path $path -Directory -Recurse
    } else {
        return @()
    }
}

function Invoke-CheckToolPresence {
    param (
        [string]$toolName,
        [string]$toolLabel
    )

    if (-not (Get-Command $toolName -ErrorAction SilentlyContinue)) {
        Write-Host "$toolLabel is required for this operation. Please install it, ensure it's in PATH and try again."
        return $false
    }
    return $true
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
                $property = switch ($mediaType) {
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
                    $firstCharacter = $_.Name.Substring(0, 1)
                    $sanitizedCharacter = $firstCharacter -replace '[\\/:*?"<>|]', "_"
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
                        if (-not (Invoke-CheckToolPresence -toolName "magick" -toolLabel "ImageMagick")) {
                            return
                        }

                        Get-ChildItem -Path $folderPath -File | ForEach-Object {
                            if ($_.Extension -match "\.(jpg|jpeg|png|gif)$") {
                                $imageInfo = & magick identify -format "%wx%h" $_.FullName
                                if ($imageInfo -match "(\d+)x(\d+)") {
                                    $width = [int]$matches[1]
                                    $height = [int]$matches[2]

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
                        if (-not (Invoke-CheckToolPresence -toolName "ffmpeg" -toolLabel "FFmpeg")) {
                            return
                        }

                        Get-ChildItem -Path $folderPath -File | ForEach-Object {
                            if ($_.Extension -match "\.(mp3|mp4|wav|avi)$") {
                                $durationInfo = & ffmpeg -i $_.FullName 2>&1 | Select-String "Duration"
                                if ($durationInfo -match "Duration: (\d+):(\d+):(\d+)") {
                                    $hours = [int]$matches[1]
                                    $minutes = [int]$matches[2]
                                    $seconds = [int]$matches[3]
                                    $totalSeconds = ($hours * 3600) + ($minutes * 60) + $seconds

                                    $durationFolder = if ($totalSeconds -le 300) {
                                        "Short (0-5 min)"
                                    } elseif ($totalSeconds -le 1800) {
                                        "Medium (5-30 min)"
                                    } elseif ($totalSeconds -le 3600) {
                                        "Long (30-60 min)"
                                    } else {
                                        "Very Long (60+ min)"
                                    }

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

    $grandOrgType = Get-ValidInput -prompt "By Date (1), By Type (2), By Size (3), By Name (4), Multimedia advanced (5), Abort (6)" -validValues @("1", "2", "3", "4", "5", "6")

    if ($grandOrgType -eq "6") {
        return
    }

    $mediaType = $null
    if ($grandOrgType -eq "1") { # Organizing by date
        $mediaType = Get-ValidInput -prompt "Creation Date (1), Last Modified Date (2), Last Accessed Date (3)" -validValues @("1", "2", "3")
    } elseif ($grandOrgType -eq "5") { # Multimedia advanced
        $mediaType = Get-ValidInput -prompt "Resolution (Images) (1), Duration (Audio/Video) (2)" -validValues @("1", "2")
        $moveToOther = Get-ValidInput -prompt "Invalid files are possible. Move to Other? Yes (Y), No (N)" -validValues @("Y", "N")
    }

    $originalSubfolders = Get-Subfolders -path $path

    # Process the main folder
    Invoke-OrganizeInFolder -folderPath $path -grandOrgType $grandOrgType -mediaType $mediaType

    $originalSubfolders | ForEach-Object {
        Invoke-OrganizeInFolder -folderPath $_.FullName -grandOrgType $grandOrgType -mediaType $mediaType
    }

    Write-Host "Files in $path have been organized."
}

function Invoke-ConvertFiles {
    param (
        [string]$path
    )

    # The true hardworker here
    function Invoke-ConvertInFolder {
        param (
            [string]$folderPath
        )

        # Abort if files with the same name exist, extension does not matter
        $existingFiles = Get-ChildItem -Path $folderPath -File | Where-Object { $_.Name -eq $_.Name }
        if ($existingFiles) {
            Write-Host "Files with the same name already exist in $folderPath. Aborting conversion."
        }

        switch ($grandConvType) {
            "1" { # Images
                $targetExtension = switch ($mediaType) {
                    "1" { ".jpg" }
                    "2" { ".jpeg" }
                    "3" { ".png" }
                    "4" { ".gif" }
                    "5" { ".webp" }
                    "6" { ".ico" }
                    "7" { ".pdf" }
                }
    
                Get-ChildItem -Path $folderPath -File | ForEach-Object {
                    if ($_.Extension -ieq $targetExtension) {
                        return
                    }
    
                    if ($_.Extension -match "\.(jpg|jpeg|png|gif|webp|ico)$") {
                        $targetFileName = [System.IO.Path]::ChangeExtension($_.FullName, $targetExtension)
                        & magick $_.FullName $targetFileName
                        if (Test-Path -Path $targetFileName) {
                            if ($EraseOriginals -eq "2") {
                                Remove-Item $_.FullName -Force
                            } else {
                                Invoke-PromptAndMoveToOther -filePath $_.FullName -folderPath $folderPath -actionType "Originals"
                            }
                        } else {
                            Write-Host "Conversion failed for file: $($_.FullName)."
                            Invoke-PromptAndMoveToOther -filePath $_.FullName -folderPath $folderPath -actionType "Failed"
                        }
                    } else {
                        if ($moveToOther -eq "Y") {
                            Invoke-PromptAndMoveToOther -filePath $_.FullName -folderPath $folderPath
                        }
                    }
                }
            }
            "2" { # Music
                $targetExtension = switch ($mediaType) {
                    "1" { ".mp3" }
                    "2" { ".wav" }
                    "3" { ".m4a" }
                    "4" { ".ogg" }
                }
    
                Get-ChildItem -Path $folderPath -File | ForEach-Object {
                    if ($_.Extension -ieq $targetExtension) {
                        return
                    }
    
                    if ($_.Extension -match "\.(mp3|wav|m4a|ogg)$") {
                        $targetFileName = [System.IO.Path]::ChangeExtension($_.FullName, $targetExtension)
                        & ffmpeg -i $_.FullName $targetFileName
                        if (Test-Path -Path $targetFileName) {
                            if ($EraseOriginals -eq "2") {
                                Remove-Item $_.FullName -Force
                            } else {
                                Invoke-PromptAndMoveToOther -filePath $_.FullName -folderPath $folderPath -actionType "Originals"
                            }
                        } else {
                            Write-Host "Conversion failed for file: $($_.FullName)."
                            Invoke-PromptAndMoveToOther -filePath $_.FullName -folderPath $folderPath -actionType "Failed"
                        }
                    } else {
                        if ($moveToOther -eq "Y") {
                            Invoke-PromptAndMoveToOther -filePath $_.FullName -folderPath $folderPath
                        }
                    }
                }
            }
            "3" { # Videos
                $targetExtension = switch ($mediaType) {
                    "1" { ".mp4" }
                    "2" { ".avi" }
                    "3" { ".mkv" }
                    "4" { ".mov" }
                }
    
                Get-ChildItem -Path $folderPath -File | ForEach-Object {
                    if ($_.Extension -ieq $targetExtension) {
                        return
                    }
    
                    if ($_.Extension -match "\.(mp4|avi|mkv|mov)$") {
                        $targetFileName = [System.IO.Path]::ChangeExtension($_.FullName, $targetExtension)
                        & ffmpeg -i $_.FullName $targetFileName
                        if (Test-Path -Path $targetFileName) {
                            if ($EraseOriginals -eq "2") {
                                Remove-Item $_.FullName -Force
                            } else {
                                Invoke-PromptAndMoveToOther -filePath $_.FullName -folderPath $folderPath -actionType "Originals"
                            }
                        } else {
                            Write-Host "Conversion failed for file: $($_.FullName)."
                            Invoke-PromptAndMoveToOther -filePath $_.FullName -folderPath $folderPath -actionType "Failed"
                        }
                    } else {
                        if ($moveToOther -eq "Y") {
                            Invoke-PromptAndMoveToOther -filePath $_.FullName -folderPath $folderPath
                        }
                    }
                }
            }
            "4" { # Documents
                $targetExtension = switch ($docType) {
                    "1" { switch ($mediaType) {
                        "1" { "pdf" }
                        "2" { "docx" }
                        "3" { "odt" }
                        "4" { "txt" }
                        "5" { "html" } 
                    } }
                    "2" { switch ($mediaType) {
                        "1" { "xlsx" }
                        "2" { "ods" }
                        "3" { "csv" }
                    } }
                }
    
                Get-ChildItem -Path $folderPath -File | ForEach-Object {
                    if ($_.Extension -ieq $targetExtension) {
                        return
                    }
                    
                    if ($docType -eq "1") {
                        if ($_.Extension -match "\.(pdf|docx|doc|odt|txt|html)$") {
                            & soffice --headless --convert-to $targetExtension $_.FullName --outdir ([System.IO.Path]::GetDirectoryName($_.FullName))
                            if (Test-Path -Path $targetFileName) {
                                if ($EraseOriginals -eq "2") {
                                    Remove-Item $_.FullName -Force
                                } else {
                                    Invoke-PromptAndMoveToOther -filePath $_.FullName -folderPath $folderPath -actionType "Originals"
                                }
                            } else {
                                Write-Host "Conversion failed for file: $($_.FullName)."
                                Invoke-PromptAndMoveToOther -filePath $_.FullName -folderPath $folderPath -actionType "Failed"
                            }
                        } else {
                            if ($moveToOther -eq "Y") {
                                Invoke-PromptAndMoveToOther -filePath $_.FullName -folderPath $folderPath
                            }
                        }
                    } else {
                        if ($_.Extension -match "\.(xlsx|ods|csv)$") {
                            & soffice --headless --convert-to $targetExtension $_.FullName --outdir ([System.IO.Path]::GetDirectoryName($_.FullName))
                            if (Test-Path -Path $targetFileName) {
                                if ($EraseOriginals -eq "2") {
                                    Remove-Item $_.FullName -Force
                                } else {
                                    Invoke-PromptAndMoveToOther -filePath $_.FullName -folderPath $folderPath -actionType "Originals"
                                }
                            } else {
                                Write-Host "Conversion failed for file: $($_.FullName)."
                                Invoke-PromptAndMoveToOther -filePath $_.FullName -folderPath $folderPath -actionType "Failed"
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

    $grandConvType = Get-ValidInput -prompt "What are we converting? Images (1), Music (2), Videos (3), Documents (4), Abort (5)" -validValues @("1", "2", "3", "4", "5")

    if ($grandConvType -eq "5") {
        return
    }

    $mediaType = $null
    switch ($grandConvType) {
        "1" { # Images
            # Check if ImageMagick is installed
            if (-not (Invoke-CheckToolPresence -toolName "magick" -toolLabel "ImageMagick")) {
                return
            }
            $mediaType = Get-ValidInput -prompt "Convert to JPEG (1), JPG (2), PNG (3), GIF (4), WebP (5), ICO (6), PDF (7)" -validValues @("1", "2", "3", "4", "5", "6", "7")
        }
        "2" { # Music
            # Check if ffmpeg is installed
            if (-not (Invoke-CheckToolPresence -toolName "ffmpeg" -toolLabel "FFmpeg")) {
                return
            }
            $mediaType = Get-ValidInput -prompt "Convert to MP3 (1), WAV (2), M4A (3), OGG (4)" -validValues @("1", "2", "3", "4")
        }
        "3" { # Videos
            # Check if ffmpeg is installed
            if (-not (Invoke-CheckToolPresence -toolName "ffmpeg" -toolLabel "FFmpeg")) {
                return
            }
            $mediaType = Get-ValidInput -prompt "Convert to MP4 (1), AVI (2), MKV (3), MOV (4)" -validValues @("1", "2", "3", "4")
        }
        "4" { # Text Documents
            # Check if LibreOffice is installed
            if (-not (Invoke-CheckToolPresence -toolName "soffice" -toolLabel "LibreOffice")) {
                return
            }
            # Ask if we're working with text or table documents
            $docType = Get-ValidInput -prompt "Type of document - Text (1), Table (2)" -validValues @("1", "2")
            if ($docType -eq "1") {
                $mediaType = Get-ValidInput -prompt "Convert to PDF (1), DOCX (2), ODT (3), TXT (4), HTML (5)" -validValues @("1", "2", "3", "4", "5")
            } else {
                $mediaType = Get-ValidInput -prompt "Convert to XLSX (1), ODS (2), CSV (3)" -validValues @("1", "2", "3")
            }
        }
    }
    $MoveToOther = Get-ValidInput -prompt "Invalid files are possible. Move to Other? Yes (Y), No (N)" -validValues @("Y", "N")

    $EraseOriginals = Get-ValidInput -prompt "Move or erase original files? Move to Originals folder (1), Erase (2) - caution advised" -validValues @("1", "2")

    $originalSubfolders = Get-Subfolders -path $path

    # Process the main folder
    Invoke-ConvertInFolder -folderPath $path -grandConvType $grandConvType -mediaType $mediaType

    $originalSubfolders | ForEach-Object {
        Invoke-ConvertInFolder -folderPath $_.FullName -grandConvType $grandConvType -mediaType $mediaType
    }

    Write-Host "Files in $path have been converted."
    
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
