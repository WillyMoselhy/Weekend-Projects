<#
.SYNOPSIS
A script to copy files from the phone camera to local disk. Uses shell to access the MTP files.

.DESCRIPTION
This is a weekend PowerShell script to help backup photos on the phone to the computer.
Inspired from https://blog.daiyanyingyu.uk/2018/03/20/powershell-mtp/

.EXAMPLE
 \\emeacssdfs.europe.corp.microsoft.com\netpod\rfl\Compare-PStatSummary.ps1 -SDPPath \\MyPC\temp\SDPs\ClusterReports
 This command will compare all !PStatSum_*.TXT in folder \\MyPC\temp\SDPs\ClusterReports

.EXAMPLE
.\Copy-MTPCameraByMonth.ps1 -MTPSourcePath WillyMobile\Card\DCIM\Camera -TargetPath D:\PhoneBackup\Camera
 This will copy files from the camera folder on the phone to the target path and segregate by month.
 
.LINK
https://plusontech.com/wp-admin/post.php?post=182&action=edit
https://github.com/WillyMoselhy/Weekend-Projects
#>

Param(
    # Folder path on the MTP device. e.g. WillyMobile\Card\DCIM\Camera
    [Parameter(Mandatory = $true)]
    [string] $MTPSourcePath,
    
    # Path to location to copy the files from MTP device
    [Parameter(Mandatory = $true)]
    [String] $TargetPath
)

function GetMTPFolder ($MTPSourcePath){
    # Synopsis Loop through path to get the desired folder

    #Get the MTP folder item
    $PathArray = $MTPSourcePath -split "\\" # Double \\ for regex escape \

    $MTPFolder = $null
    foreach($item in $PathArray){
        if(!($MTPFolder)){ # We are at the first cycle
            $MTPFolder = $Script:ShellItem.GetFolder.Items() | Where-Object{$_.Name -eq $item}
        }
        else{ #We are getting subfolders
            $MTPFolder = $MTPFolder.GetFolder.Items() | Where-Object {$_.Name -eq $item}
        } 
    } 
    return $MTPFolder
}


#Create a shell application  
$Shell = New-Object -ComObject Shell.Application

#Get the my computer list of items 
# 17 (0x11) = ssfDRIVES from the ShellSpecialFolderConstants (https://msdn.microsoft.com/en-us/library/windows/desktop/bb774096(v=vs.85).aspx)
$ShellItem = $Shell.NameSpace(17).Self

# Get the folder of the Camera using the supplied source path
$CameraFolder = GetMTPFolder -MTPSourcePath $MTPSourcePath

#Get list of images and videos in the folder
$CameraItems = $CameraFolder.GetFolder.Items()

#Get target path shell item
$TargetFolder = Get-Item -Path $TargetPath

### Copy items from camera by month
# We use the file name to arrange folders 
# File names follow this pattern yyyyMMdd_HHmmss e.g. 20200104_231922.jpg
# Files that do not match this pattern are excluded and reported

$FileNameRegex = "^(?<Year>\d{4})(?<Month>\d{2})(?<Day>\d{2})_(?<Hour>\d{2})(?<Minute>\d{2})(?<Second>\d{2}).*\.(?<Extension>.+)$" # https://regexr.com/45sdj

$ProgressActivityName = "Copying files from '$MTPSourcePath' to '$TargetPath'"

$SkippedFiles = @() 
$CopiedFilesCount = 0

foreach ($File in ($CameraItems |Sort-Object -Property Name) ){
    #Validate file name matches pattern
    
    Write-Progress -Activity $ProgressActivityName -Status "Working on it" -CurrentOperation "Copying: $($File.Name) - Finished $CopiedFilesCount / $($CameraItems.count)" -PercentComplete (($CopiedFilesCount/$CameraItems.count)*100)
    $CopiedFilesCount++

    if($File.Name -notmatch $FileNameRegex){
        $SkippedFiles += [PSCustomObject]@{
            Name       = $File.Name
            TargetPath = $null
            Reason     = "Pattern mismatch"
        }
        Write-Warning "$($File.Name) is skipped because of pattern"
    }

    else{
        $YearMonth = "{0}{1}" -f $Matches.Year,$Matches.Month
        $YearMonthFolder = New-Item -Path "$($TargetFolder.FullName)\$YearMonth" -ItemType Directory -Force

        $TargetFilePath = Join-Path -Path $YearMonthFolder -ChildPath $File.Name
        if(Test-Path -Path $TargetFilePath){ # A file with the same name already exists
            $SkippedFiles += [PSCustomObject]@{
                Name       = $File.Name
                TargetPath = $TargetFilePath
                Reason     = "Duplicate file name"
            }
            Write-Warning "$($File.Name) is skipped due to duplicate file name"
        }
        else{
            # >>>>  This is where the magic happens! <<<< #
            $TargetFolderShell = $Shell.NameSpace($YearMonthFolder.FullName).self
            
            $TargetFolderShell.GetFolder.CopyHere($File)
        }        
    } 
}

#Here is a nice view of the files that were not copied!
$SkippedFiles | Out-GridView
