﻿<#
 Start SoftPaq Downloads
 Copyright © 2015-2018 Michael 'Tex' Hex 
 Licensed under the Apache License, Version 2.0. 

 https://github.com/texhex/BiosSledgehammer
#>

#Script version
$scriptversion = "1.1.4"

#This script requires PowerShell 4.0 or higher 
#requires -version 4.0

#Require full level Administrator
#requires -runasadministrator

#Guard against common code errors
Set-StrictMode -version 2.0

#Terminate script on errors 
$ErrorActionPreference = 'Stop'

#Import Module with some helper functions
Import-Module $PSScriptRoot\MPSXM.psm1 -Force



function Get-UserConfirm()
{
    param(
        [Parameter(Mandatory = $True, ValueFromPipeline = $True)]
        [ValidateNotNullOrEmpty()]
        [string]$BaseFolder
    )

    #from https://social.technet.microsoft.com/Forums/scriptcenter/en-US/3d8f242b-199b-4d4c-b973-0246ce1c065c/windows-powershell-tip-of-the-week-is-there-an-easy-way-to-display-and-process-confirmation?forum=ITCG
    #by Shay Levi (https://social.technet.microsoft.com/profile/shay%20levi)

    $caption = "BIOS Sledgehammer: Start SoftPaq Downloads v$scriptversion"
    $message = "This script will download firmware update files from HP.com,`nbased on the SPDownload.txt files found in [$BaseFolder].`nStart downloads?"

    $yes = new-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Start downloads"
    $no = new-Object System.Management.Automation.Host.ChoiceDescription "&No", "Do not start, stop script"

    $choices = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
    $answer = $host.ui.PromptForChoice($caption, $message, $choices, 1)

    switch ($answer)
    {
        0 
        {
            return $true
            break
        }

        1 
        {
            return $false
            break
        }
    }
}


function Test-FolderStructure()
{
    param(
        [Parameter(Mandatory = $True, ValueFromPipeline = $True)]
        [ValidateNotNullOrEmpty()]
        [string]$SearchPath
    )
    if ( -not (Test-DirectoryExists "$SearchPath\Models") )
    {
        throw New-Exception -DirectoryNotFound "Path [$SearchPath\Models] does not exist"
    }

    if ( -not (Test-DirectoryExists "$SearchPath\Shared") )
    {
        throw New-Exception -DirectoryNotFound "Path [$SearchPath\Shared] does not exist"
    }

    if ( -not (Test-DirectoryExists "$SearchPath\PwdFiles") )
    {
        throw New-Exception -DirectoryNotFound "Path [$SearchPath\PwdFiles] does not exist"
    }

    if ( -not (Test-FileExists "$SearchPath\BiosSledgehammer.ps1") )
    {
        throw New-Exception -FileNotFound "File [$SearchPath\BiosSledgehammer.ps1] does not exist"
    }

}

function Remove-FileIfExists()
{
    param(
        [Parameter(Mandatory = $True, ValueFromPipeline = $True)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )
    if ( Test-FileExists -Path $Path )
    {
        try
        {
            #we allow the system half a second...
            Start-Sleep -Milliseconds 500

            Remove-Item -Path "$Path" -Force
        }
        catch
        {
            #Sometimes this does not work because AV or backup software is to eager to scan the file
            write-warning "Unable to delete file [$Path] - $($Error[0])"
        }
    }
}


function Start-DownloadFile()
{
    param(
        [Parameter(Mandatory = $True, ValueFromPipeline = $True)]
        [ValidateNotNullOrEmpty()]
        [string]$URL,

        [Parameter(Mandatory = $True, ValueFromPipeline = $True)]
        [ValidateNotNullOrEmpty()]
        [string]$DownloadPath
    )
    $file = Get-FileName($URL)

    #Ensure the files does not exist
    $tempFile = "$DownloadPath\$file"
    Remove-FileIfExists $tempFile

    $webClient = New-Object "Net.WebClient"

    write-host "  Downloading from [$URL]"
    write-host "                to [$tempFile]... " -NoNewline

    $webClient.DownloadFile($URL, $tempFile)

    $webClient = $null

    write-host "Done"

    return $tempFile
}


function Invoke-AcquireHPFile()
{
    param(
        [Parameter (Mandatory = $true)]
        [ValidateSet("SoftPaq", "ReleaseNotes")]
        [string]$Type,

        [Parameter(Mandatory = $True)]
        [ValidateNotNullOrEmpty()]
        [string]$URL,

        [Parameter(Mandatory = $True)]
        [ValidateNotNullOrEmpty()]
        [string]$DestinationPath,

        [Parameter(Mandatory = $True)]
        [ValidateNotNullOrEmpty()]
        [string]$DownloadPath
    )

    if ( $Type -eq "SoftPaq" )
    {
        write-host "  SoftPaq URL: $URL" 
    }
    else
    {
        write-host "  Release Notes URL: $URL" 
    }

    $destFilename = "$DestinationPath\$(Get-FileName($URL))"

    if ( Test-FileExists -Path $destFilename )
    {
        write-host "  File already exists"
    }
    else
    {
        $tempFile = Start-DownloadFile -URL $URL -DownloadPath $DownloadPath

        if ( $Type -eq "SoftPaq" )
        {
            $SPName = Get-FileName -Path $tempFile -WithoutExtension
            $SPName = $SPName.ToUpper()

            write-host "  Extracting SoftPaq... " -NoNewline
            &$tempFile -e -s     
            Start-Sleep -Seconds 2 #just to be sure, sometimes it requires some extra time
            write-host "Done"

            #We are expecting the file to be present in C:\SWSetup\SPxxxx
            $SPExtractedPath = "$UNPACK_FOLDER\$SPName"

            #SP88497 is extracted to [SP88497, so we better make sure to check that folder as well
            $SPExtractedPath2 = "$UNPACK_FOLDER\[$($SPName)"
            if ( Test-Path -LiteralPath $SPExtractedPath2 -PathType Container) 
            {
                write-warning "Extracted files folder is $SPExtractedPath2, renaming folder"
                Rename-Item -LiteralPath $SPExtractedPath2 -NewName $SPName
            }
          
            if ( -not (Test-DirectoryExists -Path $SPExtractedPath) )            
            {
                throw New-Exception -FileNotFound "Unable to locate unpack folder [$SPExtractedPath]"
            }
            else
            {
                #now we need to copy the unpacked files
                write-host "  Copy from [$SPExtractedPath] to [$destFolder]... " -NoNewline
                $ignored = Get-ChildItem -Path $SPExtractedPath | Copy-Item -Destination $destFolder -Recurse -Container -Force
                write-host "Done"

                #Remove extract folder - this will sometimes fail because of backup or AV tools
                try 
                {
                    $ignored = Remove-Item -LiteralPath $SPExtractedPath -Recurse -Force
                }
                catch
                {
                    write-warning "Unable to remove temp extraction folder [$SPExtractedPath] - $($Error[0])"
                }
            }
        }

        #copy SPXXXX to destination        
        Copy-FileToDirectory -Filename $tempFile -Directory $DestinationPath

        #This would be a good idea in case the SPxxx would not be running sometimes in the background
        #and make this call fail. Hence: Skip it. 
        #Remove-FileIfExists -Path $tempFile

        write-host "  File processed successfully"
    }
}


function Invoke-DownloadProcess()
{
    param(
        [Parameter(Mandatory = $True, ValueFromPipeline = $True)]
        [ValidateNotNullOrEmpty()]
        [string]$SettingsFile,

        [Parameter(Mandatory = $True, ValueFromPipeline = $True)]
        [ValidateNotNullOrEmpty()]
        [string]$DownloadPath
    )

    $destFolder = Get-ContainingDirectory($SettingsFile)

    $settings = Read-StringHashtable $SettingsFile

    if ($settings.ContainsKey("NoteURL"))
    {        
        $URL = $settings["NoteURL"]
        $type = "ReleaseNotes"

        Invoke-AcquireHPFile -Type $type -URL $URL -DestinationPath $destFolder -DownloadPath $DownloadPath
    }

    if ($settings.ContainsKey("SPaqURL"))
    {
        $URL = $settings["SPaqURL"]
        $type = "SoftPaq"

        Invoke-AcquireHPFile -Type $type -URL $URL -DestinationPath $destFolder -DownloadPath $DownloadPath
    }

}

#Issue #41 - https://github.com/texhex/BiosSledgehammer/issues/41
#
#TLS 1.0 and SSL3 are no longer supported by ftp.hp.com (see http://ssl-checker.online-domain-tools.com/
#and enter ftp.hp.com)
#
#The default enabled security protocols for .NET 4.0/4.5 (and hence PowerShell) are SecurityProtocolType.Tls|SecurityProtocolType.Ssl3.
#
#Therefore we need to turn on TLS 1.1 and TLS 1.2. As TLS 1.3 is also on the horizon, and will properly 
#be a default in PowerShell, we only turn on TLS 1.1 and TLS 1.2 without touching the default protocols. 
#
#Full details on this StackOverflow answer by Luke Hutton: https://stackoverflow.com/a/28333370
#
Set-HTTPSecurityProtocolSecureDefault


######################################################
## Main ##############################################
######################################################


Set-Variable CHECK_FILENAME "$PSScriptRoot\BiosSledgehammer.ps1" –option ReadOnly -Force
Set-Variable TEMP_DOWNLOAD_FOLDER "$(Get-TempFolder)\TempDownload" –option ReadOnly -Force
Set-Variable UNPACK_FOLDER "C:\SWSetup" –option ReadOnly -Force


if ( Get-UserConfirm -BaseFolder $PSScriptRoot)
{
    $ignored = Test-FolderStructure -SearchPath $PSScriptRoot

    #ensure the temp downloads folder exists
    $ignored = New-Item -Path $TEMP_DOWNLOAD_FOLDER -ItemType Directory -Force

    #scan for SPDownload.txt files
    $Files = Get-ChildItem -Path $PSScriptRoot -Filter "SPDownload.txt" -Recurse
    foreach ($file in $Files)
    { 
        $curFile = $file.Fullname
        write-host "File [$curFile]"
 
        Invoke-DownloadProcess -SettingsFile $curFile -DownloadPath $TEMP_DOWNLOAD_FOLDER  
    }

    write-host "All done, waiting 20 seconds before starting clean up..."
    Start-Sleep -Seconds 20


    write-host "Cleaning up $UNPACK_FOLDER..."
    #check if the UNPACK_FOLDER exists and if it's empty. If so, delete it
    if ( Test-DirectoryExists $UNPACK_FOLDER )
    {
        $count = (Get-ChildItem -Path $UNPACK_FOLDER -Directory | Measure-Object).Count
       
        if ( $count -eq 0 )
        {
            #Folder is empty
            $ignored = Remove-Item -Path $UNPACK_FOLDER -Force
        }
    }

    write-host "Cleaning up $TEMP_DOWNLOAD_FOLDER..."
    #try to delete the temp folder
    try
    {
        $ignored = Remove-Item -LiteralPath $TEMP_DOWNLOAD_FOLDER -Force -Recurse -Confirm:$false
    }
    catch
    {
        Write-Warning "Unable to delete folder [$TEMP_DOWNLOAD_FOLDER] - $($Error[0])"
    }
}

write-host "Script finished!"

