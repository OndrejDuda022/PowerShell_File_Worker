# DEFINE PARAMETERS
param (
    [string]$fPath,
    [string]$act,
    [string]$grandOrgType,
    [string]$mediaType,
    [string]$moveToOther,
    [string]$grandConvType,
    [string]$docType,
    [string]$eraseOriginals
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

function Get-UniqueFileName {
    param (
        [string]$filePath,
        [string]$targetExtension,
        [string]$jobType
    )

    $directory = [System.IO.Path]::GetDirectoryName($filePath)
    $fileNameWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($filePath)
    $originalExtension = [System.IO.Path]::GetExtension($filePath)

    $uniqueFilePath = Join-Path -Path $directory -ChildPath "$fileNameWithoutExtension$targetExtension"
    $counter = 1

    while (Test-Path -Path $uniqueFilePath) {
        $uniqueFilePath = Join-Path -Path $directory -ChildPath "$fileNameWithoutExtension ($counter)$targetExtension"
        $counter++
    }

    if ($jobtype -eq "docs") {
        $uniqueFilePath = Join-Path -Path $directory -ChildPath "$fileNameWithoutExtension ($($counter - 1))$originalExtension"
    }
    else {
        $uniqueFilePath = Join-Path -Path $directory -ChildPath "$fileNameWithoutExtension ($($counter - 1))$targetExtension"
    }
    

    return $uniqueFilePath
}

function Invoke-ValidateParameter {
    param (
        [string]$toCheck,
        [string]$grandOrgType,
        [string]$mediaType,
        [string]$moveToOther,
        [string]$grandConvType,
        [string]$docType,
        [string]$eraseOriginals
    )
    switch ($toCheck) {
        "grandOrgType" {
            if ($grandOrgType -and $grandOrgType -notin @("1", "2", "3", "4", "5", "6")) {
                Write-Host "Error: Invalid parameter - grandOrgType."
                return $false
            }
        }
        "mediaType" {
            if ($mediaType -and $grandOrgType -eq "1" -and $mediaType -notin @("1", "2", "3")) {
                Write-Host "Error: Invalid parameter - mediaType."
                return $false
            }
            if ($mediaType -and $grandOrgType -eq "5" -and $mediaType -notin @("1", "2")) {
                Write-Host "Error: Invalid parameter - mediaType."
                return $false
            }
            if ($mediaType -and $grandConvType -eq "1" -and $mediaType -notin @("1", "2", "3", "4", "5", "6", "7")) {
                Write-Host "Error: Invalid parameter - mediaType."
                return $false
            }
            if ($mediaType -and $grandConvType -eq "2" -and $mediaType -notin @("1", "2", "3", "4")) {
                Write-Host "Error: Invalid parameter - mediaType."
                return $false
            }
            if ($mediaType -and $grandConvType -eq "3" -and $mediaType -notin @("1", "2", "3", "4")) {
                Write-Host "Error: Invalid parameter - mediaType."
                return $false
            }
        }
        "grandConvType" {
            if ($grandConvType -and $grandConvType -notin @("1", "2", "3", "4", "5")) {
                Write-Host "Error: Invalid parameter - grandConvType."
                return $false
            }
        }
        "docType" {
            if ($docType -and $docType -notin @("1", "2")) {
                Write-Host "Error: Invalid parameter - docType."
                return $false
            }
            if ($docType -and $grandConvType -eq "4" -and $docType -eq "1" -and $mediaType -notin @("1", "2", "3", "4", "5")) {
                Write-Host "Error: Invalid parameter - mediaType."
                return $false
            }
            if ($docType -and $grandConvType -eq "4" -and $docType -eq "2" -and $mediaType -notin @("1", "2", "3")) {
                Write-Host "Error: Invalid parameter - mediaType."
                return $false
            }
        }
        "moveToOther" {
            if ($moveToOther -and $moveToOther -notin @("Y", "N")) {
                Write-Host "Error: Invalid parameter - moveToOther."
                return $false
            }
        }
        "eraseOriginals" {
            if ($eraseOriginals -and $eraseOriginals -notin @("1", "2")) {
                Write-Host "Error: Invalid parameter - eraseOriginals."
                return $false
            }
        }
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
        [string]$path,
        [string]$grandOrgType,
        [string]$mediaType,
        [string]$moveToOther
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

    $originalSubfolders = Get-Subfolders -path $path

    Invoke-OrganizeInFolder -folderPath $path -grandOrgType $grandOrgType -mediaType $mediaType

    $originalSubfolders | ForEach-Object {
        Invoke-OrganizeInFolder -folderPath $_.FullName -grandOrgType $grandOrgType -mediaType $mediaType
    }

    Write-Host "Organizing of files in $path is complete."
}

function Invoke-ConvertFiles {
    param (
        [string]$path,
        [string]$grandConvType,
        [string]$mediaType,
        [string]$moveToOther,
        [string]$docType,
        [string]$eraseOriginals
    )

    # The true hardworker here
    function Invoke-ConvertInFolder {
        param (
            [string]$folderPath
        )

        switch ($grandConvType) {
            "1" { # Images
                if (-not (Invoke-CheckToolPresence -toolName "magick" -toolLabel "ImageMagick")) {
                    return
                }

                $targetExtension = switch ($mediaType) {
                    "1" { ".jpeg" }
                    "2" { ".jpg" }
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
    
                    if ($_.Extension -match "\.(jpg|jpeg|png|webp|ico)$") {
                        $targetFileName = [System.IO.Path]::ChangeExtension($_.FullName, $targetExtension)

                        if (Test-Path -Path $targetFileName) {
                            $targetFileName = Get-UniqueFileName -filePath $_.FullName -targetExtension $targetExtension -jobType "images"
                        }

                        & magick $_.FullName $targetFileName
                        if (Test-Path -Path $targetFileName) {
                            if ($eraseOriginals -eq "2") {
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
                if (-not (Invoke-CheckToolPresence -toolName "ffmpeg" -toolLabel "FFmpeg")) {
                    return
                }
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

                        if (Test-Path -Path $targetFileName) {
                            $targetFileName = Get-UniqueFileName -filePath $_.FullName -targetExtension $targetExtension -jobType "audio"
                        }

                        & ffmpeg -i $_.FullName $targetFileName
                        if (Test-Path -Path $targetFileName) {
                            if ($eraseOriginals -eq "2") {
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
                if (-not (Invoke-CheckToolPresence -toolName "ffmpeg" -toolLabel "FFmpeg")) {
                    return
                }
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

                        if (Test-Path -Path $targetFileName) {
                            $targetFileName = Get-UniqueFileName -filePath $_.FullName -targetExtension $targetExtension -jobType "video"
                        }

                        & ffmpeg -i $_.FullName $targetFileName
                        if (Test-Path -Path $targetFileName) {
                            if ($eraseOriginals -eq "2") {
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
                if (-not (Invoke-CheckToolPresence -toolName "soffice" -toolLabel "LibreOffice")) {
                    return
                }
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
                    $dotTargetExtension = ".$targetExtension"
                    if ($_.Extension -ieq $dotTargetExtension) {
                        return
                    }
                    
                    if ($docType -eq "1") {
                        if ($_.Extension -match "\.(docx|doc|odt|txt|html)$") {
                            $targetFileName = [System.IO.Path]::ChangeExtension($_.FullName, $targetExtension)

                            if (Test-Path -Path $targetFileName) {
                                $uniqueTargetFileName = Get-UniqueFileName -filePath $_.FullName -targetExtension $dotTargetExtension -jobType "docs"
                                Copy-Item -Path $_.FullName -Destination $uniqueTargetFileName
                                & soffice --headless --convert-to $targetExtension $uniqueTargetFileName --outdir ([System.IO.Path]::GetDirectoryName($_.FullName))
                                Remove-Item $uniqueTargetFileName -Force

                            } else {
                                & soffice --headless --convert-to $targetExtension $_.FullName --outdir ([System.IO.Path]::GetDirectoryName($_.FullName))
                            }

                            if (Test-Path -Path $targetFileName) {
                                if ($eraseOriginals -eq "2") {
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
                            $targetFileName = [System.IO.Path]::ChangeExtension($_.FullName, $targetExtension)

                            if (Test-Path -Path $targetFileName) {
                                $uniqueTargetFileName = Get-UniqueFileName -filePath $_.FullName -targetExtension $dotTargetExtension -jobType "docs"
                                Copy-Item -Path $_.FullName -Destination $uniqueTargetFileName
                                & soffice --headless --convert-to $targetExtension $uniqueTargetFileName --outdir ([System.IO.Path]::GetDirectoryName($_.FullName))
                                Remove-Item $uniqueTargetFileName -Force

                            } else {
                                & soffice --headless --convert-to $targetExtension $_.FullName --outdir ([System.IO.Path]::GetDirectoryName($_.FullName))
                            }

                            if (Test-Path -Path $targetFileName) {
                                if ($eraseOriginals -eq "2") {
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

    $originalSubfolders = Get-Subfolders -path $path

    # Process the main folder
    Invoke-ConvertInFolder -folderPath $path -grandConvType $grandConvType -mediaType $mediaType

    $originalSubfolders | ForEach-Object {
        Invoke-ConvertInFolder -folderPath $_.FullName -grandConvType $grandConvType -mediaType $mediaType
    }

    Write-Host "Conversion of files in $path is complete."
    
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

# Perform the action - add parameters
do {
    switch ($act) {
        "1" {
            if (-not $grandOrgType) {
                $grandOrgType = Get-ValidInput -prompt "By Date (1), By Type (2), By Size (3), By Name (4), Multimedia advanced (5), Abort (6)" -validValues @("1", "2", "3", "4", "5", "6")
                if ($grandOrgType -eq "6") {
                    $grandOrgType = $null
                    $mediaType = $null
                    $moveToOther = $null
                    $act = Select-Action
                    Continue
                }
            } else {
                $safetyCheck = Invoke-ValidateParameter -toCheck "grandOrgType" -grandOrgType $grandOrgType
                if (-not $safetyCheck) {
                    $grandOrgType = $null
                    Continue
                }
            }
            if (-not $mediaType) {
                if ($grandOrgType -eq "1") { # Organizing by date
                    $mediaType = Get-ValidInput -prompt "Creation Date (1), Last Modified Date (2), Last Accessed Date (3)" -validValues @("1", "2", "3")
                } elseif ($grandOrgType -eq "5") { # Multimedia advanced
                    $mediaType = Get-ValidInput -prompt "Resolution (Images) (1), Duration (Audio/Video) (2)" -validValues @("1", "2")
                    if (-not $moveToOther) {
                        $moveToOther = Get-ValidInput -prompt "Invalid files are possible. Move to Other? Yes (Y), No (N)" -validValues @("Y", "N")
                    } else {
                        $safetyCheck = Invoke-ValidateParameter -toCheck "moveToOther" -moveToOther $moveToOther
                        if (-not $safetyCheck) {
                            $moveToOther = $null
                            Continue
                        }
                    }
                }
            } else {
                $safetyCheck = Invoke-ValidateParameter -toCheck "mediaType" -grandOrgType $grandOrgType -mediaType $mediaType
                if (-not $safetyCheck) {
                    $mediaType = $null
                    Continue
                }
            }

            Invoke-OrganizeFiles -path $fPath -grandOrgType $grandOrgType -mediaType $mediaType -moveToOther $moveToOther

            $grandOrgType = $null
            $mediaType = $null
            $moveToOther = $null
            $act = Select-Action
        }
        "2" {
            if (-not $grandConvType) {
                $grandConvType = Get-ValidInput -prompt "What are we converting? Images (1), Music (2), Videos (3), Documents (4), Abort (5)" -validValues @("1", "2", "3", "4", "5")
                if ($grandConvType -eq "5") {
                    $grandConvType = $null
                    $mediaType = $null
                    $docType = $null
                    $moveToOther = $null
                    $eraseOriginals = $null
                    $act = Select-Action
                    Continue
                }
            } else {
                $safetyCheck = Invoke-ValidateParameter -toCheck "grandConvType" -grandConvType $grandConvType
                if (-not $safetyCheck) {
                    $grandConvType = $null
                    Continue
                }
            }
            if (-not $mediaType) {
                switch ($grandConvType) {
                    "1" { # Images
                        $mediaType = Get-ValidInput -prompt "Convert to JPEG (1), JPG (2), PNG (3), GIF (4), WebP (5), ICO (6), PDF (7)" -validValues @("1", "2", "3", "4", "5", "6", "7")
                    }
                    "2" { # Music
                        $mediaType = Get-ValidInput -prompt "Convert to MP3 (1), WAV (2), M4A (3), OGG (4)" -validValues @("1", "2", "3", "4")
                    }
                    "3" { # Videos
                        $mediaType = Get-ValidInput -prompt "Convert to MP4 (1), AVI (2), MKV (3), MOV (4)" -validValues @("1", "2", "3", "4")
                    }
                    "4" { # Text Documents
                        # Ask if we're working with text or table documents
                        if (-not $docType) {
                            $docType = Get-ValidInput -prompt "Type of document - Text (1), Table (2)" -validValues @("1", "2")
                            if ($docType -eq "1") {
                                $mediaType = Get-ValidInput -prompt "Convert to PDF (1), DOCX (2), ODT (3), TXT (4), HTML (5)" -validValues @("1", "2", "3", "4", "5")
                            } else {
                                $mediaType = Get-ValidInput -prompt "Convert to XLSX (1), ODS (2), CSV (3)" -validValues @("1", "2", "3")
                            }
                        } else {
                            $safetyCheck = Invoke-ValidateParameter -toCheck "docType" -grandConvType $grandConvType -docType $docType
                            if (-not $safetyCheck) {
                                $docType = $null
                                Continue
                            }
                        }
                    }
                }
            } else {
                $safetyCheck = Invoke-ValidateParameter -toCheck "mediaType" -grandConvType $grandConvType -mediaType $mediaType
                if (-not $safetyCheck) {
                    $mediaType = $null
                    Continue
                }
            }

            if (-not $moveToOther) {
                $moveToOther = Get-ValidInput -prompt "Invalid files are possible. Move to Other? Yes (Y), No (N)" -validValues @("Y", "N")
            } else {
                $safetyCheck = Invoke-ValidateParameter -toCheck "moveToOther" -moveToOther $moveToOther
                if (-not $safetyCheck) {
                    $moveToOther = $null
                    Continue
                }
            }

            if (-not $eraseOriginals) {
                $eraseOriginals = Get-ValidInput -prompt "Move or erase original files? Move to Originals folder (1), Erase (2) - caution advised" -validValues @("1", "2")
            } else {
                $safetyCheck = Invoke-ValidateParameter -toCheck "eraseOriginals" -eraseOriginals $eraseOriginals
                if (-not $safetyCheck) {
                    $eraseOriginals = $null
                    Continue
                }
            }

            Invoke-ConvertFiles -path $fPath -grandConvType $grandConvType -mediaType $mediaType -docType $docType -moveToOther $moveToOther -EraseOriginals $eraseOriginals

            $grandConvType = $null
            $mediaType = $null
            $docType = $null
            $moveToOther = $null
            $eraseOriginals = $null
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
