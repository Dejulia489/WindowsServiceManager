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
    } finally {
        Trace-VstsLeavingInvocation $MyInvocation
    }
}

function Get-WindowsService {
    param
    (
        $ServiceName
    )
    Get-WmiObject -Class Win32_Service | Where-Object { $PSItem.Name -eq $ServiceName }
}

function Start-WindowsService {
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

function Get-FullExecuteablePath {
    [CmdletBinding()]
    param (
        [string]
        $StringContainingPath,

        [switch]
        $JustParentPath = $false
    )
    # pattern to analyse Service Startup Command 
    $matchPattern = '( |^)(?<path>([a-zA-Z]):\\([\\\w\/.-]+)(.exe|.dll))|(( "|^")(?<path2>(([a-zA-Z]):\\([\\\w\/. -]+)(.exe|.dll)))(" |"$))'

    # check if PathName can be processed
    if($StringContainingPath -notmatch $matchPattern) {
        return Write-Error -Message "String can't be parsed. The StringContainingPath parameter should contain a valid Path ending with an '.exe' or '.dll'. Current string [$StringContainingPath]"
    }

    # extract Path
    $matchedPath = if ($matches.path) { $matches.path } else { $matches.path2 }

    if ($JustParentPath) {
        return ($matchedPath | Split-Path -Parent).Replace('"', '')
    }
    
    return $matchedPath
}