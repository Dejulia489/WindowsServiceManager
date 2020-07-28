function Add-LocalUserToLogonAsAService
{
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true)]
        [string]$user
    )
   
    PROCESS
    {
        if ( [string]::IsNullOrEmpty($user) )
        {
            return Write-Error "no account specified"
        }

        $sidstr = $null
        try
        {
            $ntprincipal = new-object System.Security.Principal.NTAccount "$user"
            $sid = $ntprincipal.Translate([System.Security.Principal.SecurityIdentifier])
            $sidstr = $sid.Value.ToString()
        }
        catch
        {
            $sidstr = $null
        }

        Write-Host "Account: $($user)" -ForegroundColor DarkCyan

        if ( [string]::IsNullOrEmpty($sidstr) )
        {
            return Write-Error "Account not found!"
        }

        Write-Host "Account SID: $($sidstr)" -ForegroundColor DarkCyan

        $tmp = [System.IO.Path]::GetTempFileName()

        Write-Host "Export current Local Security Policy" -ForegroundColor DarkCyan
        secedit.exe /export /cfg "$($tmp)" 

        $c = Get-Content -Path $tmp 

        $currentSetting = ""

        foreach ($s in $c)
        {
            if ( $s -like "SeServiceLogonRight*")
            {
                $x = $s.split("=", [System.StringSplitOptions]::RemoveEmptyEntries)
                $currentSetting = $x[1].Trim()
            }
        }

        if ( $currentSetting -notlike "*$($sidstr)*" )
        {
            Write-Host "Modify Setting ""Logon as a Service""" -ForegroundColor DarkCyan
			
            if ( [string]::IsNullOrEmpty($currentSetting) )
            {
                $currentSetting = "*$($sidstr)"
            }
            else
            {
                $currentSetting = "*$($sidstr),$($currentSetting)"
            }
			
            Write-Host "$currentSetting"
			
            $outfile = @"
[Unicode]
Unicode=yes
[Version]
signature="`$CHICAGO`$"
Revision=1
[Privilege Rights]
SeServiceLogonRight = $($currentSetting)
"@

            $tmp2 = [System.IO.Path]::GetTempFileName()
			
			
            Write-Host "Import new settings to Local Security Policy" -ForegroundColor DarkCyan
            $outfile | Set-Content -Path $tmp2 -Encoding Unicode -Force

            #notepad.exe $tmp2
            Push-Location (Split-Path $tmp2)
			
            try
            {
                secedit.exe /configure /db "secedit.sdb" /cfg "$($tmp2)" /areas USER_RIGHTS 
                #Write-Host "secedit.exe /configure /db ""secedit.sdb"" /cfg ""$($tmp2)"" /areas USER_RIGHTS "
            }
            finally
            {	
                Pop-Location
            }
        }
        else
        {
            Write-Host "NO ACTIONS REQUIRED! Account already in ""Logon as a Service""" -ForegroundColor DarkCyan
        }

        Write-Host "Done." -ForegroundColor DarkCyan
    }
}

function Get-WSMNewPSSessionOption
{
    [CmdletBinding()]
    param(
        [string] $arguments
    )
    Trace-VstsEnteringInvocation $MyInvocation
    try
    {
        $commandString = "New-PSSessionOption $arguments"
        Write-Verbose "New-PSSessionOption command: $commandString"
        return (Invoke-Expression -Command $commandString)
    }
    finally
    {
        Trace-VstsLeavingInvocation $MyInvocation
    }
}

function Get-WSMWindowsService
{
    [CmdletBinding()]
    param   (
        $ServiceName
    )
    return Get-WmiObject -Class Win32_Service | Where-Object { $PSItem.Name -eq $ServiceName }
}

function Start-WSMWindowsService
{
    [CmdletBinding()]
    param
    (
        $ServiceName
    )
    Write-Host "[$env:ComputerName]: Starting [$ServiceName]"
    $serviceObject = Get-WSMWindowsService -ServiceName $ServiceName
    $respone = $serviceObject.StartService()
    if ($respone.ReturnValue -ne 0)
    {
        return Write-Error -Message "[$env:ComputerName]: Service responded with [$($respone.ReturnValue)]. See https://docs.microsoft.com/en-us/windows/desktop/cimwin32prov/startservice-method-in-class-win32-service for details."
    }
    else
    {
        Write-Host "[$env:ComputerName]: [$ServiceName] started successfully!"
    }
}

function Stop-WSMWindowsService
{
    [CmdletBinding()]
    param
    (
        [string]
        $ServiceName,

        [int]
        $Timeout = 30,

        [switch]
        $StopProcess = $false
    )

    $serviceObject = Get-WSMWindowsService -ServiceName $ServiceName
    
    if ($serviceObject.State -eq 'Running')
    {
        $stopServiceTimer = [Diagnostics.Stopwatch]::StartNew()
        Write-Host "[$env:ComputerName]: Stopping Service [$ServiceName]"
        do
        {
            $serviceObject = Get-WSMWindowsService -ServiceName $ServiceName
            $results = $serviceObject.StopService()

            if ($stopServiceTimer.Elapsed.TotalSeconds -gt $Timeout)
            {
                if ($StopProcess)
                {
                    Write-Verbose "[$env:ComputerName]: [$ServiceName] did not respond within [$Timeout] seconds, stopping process."

                    $parentPath = Get-WSMFullExecuteablePath -StringContainingPath $serviceObject.PathName -JustParentPath

                    $allProcesses = Get-Process
                    $process = $allProcesses | Where-Object { $_.Path -like "$parentPath\*" }
                    if ($process)
                    {
                        Write-Warning -Message "[$env:ComputerName]: Files are still in use by [$($process.ProcessName)], stopping the process!"
                        $process | Stop-Process -Force -ErrorAction SilentlyContinue
                    }
                }
                else
                {
                    return Write-Error -Message "[$env:ComputerName]: [$ServiceName] did not respond within [$Timeout] seconds."             
                }
            }
            $serviceObject = Get-WSMWindowsService -ServiceName $ServiceName
        }
        while ($serviceObject.State -ne 'Stopped')

        Write-Host "[$env:ComputerName]: Stopped Service [$ServiceName]"
    }
    return $serviceObject    
}

function Get-WSMFullExecuteablePath
{
    [CmdletBinding()]
    param 
    (
        [string]
        $StringContainingPath,

        [switch]
        $JustParentPath = $false
    )
    Write-Verbose "StartCommand=[$StartCommand]"
    Write-Verbose "StringContainingPath=[$StringContainingPath]"
    # pattern to analyse Service Startup Command 
    $matchPattern = '( |^)(?<path>([a-zA-Z]):\\([\\\w\/\(\)\[\]{}öäüÖÄÜ°^!§$%&=`´,;@#+._-]+)(.exe|.dll))|(( "|^")(?<path2>(([a-zA-Z]):\\([\\\w\/\(\)\[\]{}öäüÖÄÜ°^!§$%&=`´,;@#+._ -]+)(.exe|.dll)))(" |"$))'

    # check if PathName can be processed
    #if ($StringContainingPath -notmatch $matchPattern)
    if ($StartCommand -notmatch $matchPattern) #as workaround it will not be Used the parameter. The function Parameter always trims the Quotes.
    {
        return Write-Warning -Message "String can't be parsed. The StringContainingPath parameter should contain a valid Path without spaces or wrapped in quotes and ending with an '.exe' or '.dll'. Current string [$StringContainingPath]"
    }

    # extract Path
    $matchedPath = if ($matches.path) { $matches.path } else { $matches.path2 }

    if ($JustParentPath)
    {
        return ($matchedPath | Split-Path -Parent).Replace('"', '')
    }
    
    return $matchedPath
}


function New-WSMServiceDirectory
{
    [CmdletBinding()]
    param (
        [string]
        $ParentPath
    )
    if (-not(Test-Path $ParentPath))
    {
        Write-Host "[$env:ComputerName]: Creating the service directory at [$ParentPath]."
        $null = New-Item -Path $ParentPath -ItemType 'Directory' -Force
    }  
}

function Copy-WSMServiceBinaries
{
    [CmdletBinding()]
    param (
        [string]
        $ArtifactPath,

        [string]
        $ParentPath
    )
    Write-Host "[$env:ComputerName]: Copying [$ArtifactPath] to [$ParentPath]"
    if ($ArtifactPath.EndsWith('.zip'))
    {
        Expand-Archive -Path $ArtifactPath -DestinationPath $ParentPath -Force -ErrorAction Stop
    }
    else
    {
        Copy-Item -Path "$ArtifactPath\*" -Destination $ParentPath -Force -Recurse -ErrorAction Stop 
    }
}

function Invoke-WSMCleanInstall
{
    [CmdletBinding()]
    param (
        [string]
        $ParentPath, 

        [bool]
        $StopProcess,

        [int]
        $Timeout,

        [string]
        $ServiceName
    )
    $cleanInstalltimer = [Diagnostics.Stopwatch]::StartNew()
    do
    {
        try
        {
            Get-ChildItem -Path $parentPath -Recurse -Force | Remove-Item -Recurse -Force -ErrorAction Stop
        }
        catch
        {
            switch -Wildcard ($PSItem.ErrorDetails.Message)
            {
                '*Cannot remove*'
                {
                    if ($StopProcess)
                    {
                        Write-Verbose "[$env:ComputerName]: [$ServiceName] did not respond within [$Timeout] seconds, stopping process." 
                        $allProcesses = Get-Process
                        $process = $allProcesses | Where-Object { $_.Path -like "$parentPath\*" } 
                        if ($process)
                        {
                            Write-Warning "[$env:ComputerName]: Files are still in use by [$($process.ProcessName)], stopping the process!"
                            $process | Stop-Process -Force -ErrorAction SilentlyContinue
                        }
                    }
                    else
                    {
                        return Write-Error $PSItem
                    }    
                }
                default
                {
                    return Write-Error $PSItem
                }
            }
        }
        if ($cleanInstalltimer.Elapsed.TotalSeconds -gt $Timeout)
        {
            return Write-Error "[$env:ComputerName]: [$ServiceName] did not respond within [$Timeout] seconds, clean install has failed."
        }
    }
    while (Get-ChildItem -Path $parentPath -Recurse -Force)
    $null = New-Item -ItemType Directory -Path $parentPath -Force

}