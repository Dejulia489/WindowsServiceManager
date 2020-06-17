function Get-NewPSSessionOption {
    [CmdletBinding()]
    param(
        [string] $arguments
    )
    Trace-VstsEnteringInvocation $MyInvocation
    try {
        $commandString = "New-PSSessionOption $arguments"
        Write-Verbose "New-PSSessionOption command: $commandString"
        return (Invoke-Expression -Command $commandString)
    }
    finally {
        Trace-VstsLeavingInvocation $MyInvocation
    }
}

function Get-WindowsService {
    [CmdletBinding()]
    param   (
        $ServiceName
    )
    return Get-WmiObject -Class Win32_Service | Where-Object { $PSItem.Name -eq $ServiceName }
}

function Start-WindowsService {
    [CmdletBinding()]
    param
    (
        $ServiceName
    )
    Write-Output "[$env:ComputerName]: Starting [$ServiceName]"
    $serviceObject = Get-WindowsService -ServiceName $ServiceName
    $respone = $serviceObject.StartService()
    if ($respone.ReturnValue -ne 0) {
        return Write-Error -Message "[$env:ComputerName]: Service responded with [$($respone.ReturnValue)]. See https://docs.microsoft.com/en-us/windows/desktop/cimwin32prov/startservice-method-in-class-win32-service for details."
    }
    else {
        Write-Output "[$env:ComputerName]: [$ServiceName] started successfully!"
    }
}

function Stop-WindowsService {
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

    $serviceObject = Get-WindowsService -ServiceName $ServiceName
    
    if ($serviceObject.State -eq 'Running') {
        $stopServiceTimer = [Diagnostics.Stopwatch]::StartNew()
        Write-Output "[$env:ComputerName]: Stopping Service [$ServiceName]"
        do {
            $serviceObject = Get-WindowsService -ServiceName $ServiceName
            $results = $serviceObject.StopService()

            if ($stopServiceTimer.Elapsed.TotalSeconds -gt $Timeout) {
                if ($StopProcess) {
                    Write-Verbose "[$env:ComputerName]: [$ServiceName] did not respond within [$Timeout] seconds, stopping process."

                    $parentPath = Get-FullExecuteablePath -StringContainingPath $serviceObject.PathName -JustParentPath

                    $allProcesses = Get-Process
                    $process = $allProcesses | Where-Object { $_.Path -like "$parentPath\*" }
                    if ($process) {
                        Write-Warning -Message "[$env:ComputerName]: Files are still in use by [$($process.ProcessName)], stopping the process!"
                        $process | Stop-Process -Force -ErrorAction SilentlyContinue
                    }
                }
                else {
                    return Write-Error -Message "[$env:ComputerName]: [$ServiceName] did not respond within [$Timeout] seconds."             
                }
            }
            $serviceObject = Get-WindowsService -ServiceName $ServiceName
        }
        while ($serviceObject.State -ne 'Stopped')

        Write-Output "[$env:ComputerName]: Stopped Service [$ServiceName]"
    }
    return $serviceObject    
}

function Get-FullExecuteablePath {
    [CmdletBinding()]
    param 
    (
        [string]
        $StringContainingPath,

        [switch]
        $JustParentPath = $false
    )
    # pattern to analyse Service Startup Command 
    $matchPattern = '( |^)(?<path>([a-zA-Z]):\\([\\\w\/.-]+)(.exe|.dll))|(( "|^")(?<path2>(([a-zA-Z]):\\([\\\w\/. -]+)(.exe|.dll)))(" |"$))'

    # check if PathName can be processed
    if ($StringContainingPath -notmatch $matchPattern) {
        return Write-Error -Message "String can't be parsed. The StringContainingPath parameter should contain a valid Path ending with an '.exe' or '.dll'. Current string [$StringContainingPath]"
    }

    # extract Path
    $matchedPath = if ($matches.path) { $matches.path } else { $matches.path2 }

    if ($JustParentPath) {
        return ($matchedPath | Split-Path -Parent).Replace('"', '')
    }
    
    return $matchedPath
}
