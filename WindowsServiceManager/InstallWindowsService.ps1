param
(
    [Parameter()]
    [string]
    $ServiceName = (Get-VstsInput -Name 'ServiceName' -Require), 

    [Parameter()]
    [string]
    $ArtifactPath = (Get-VstsInput -Name 'ArtifactPath' -Require),

    [Parameter()]
    $CleanInstall = (Get-VstsInput -Name 'CleanInstall' -AsBool)
)
Trace-VstsEnteringInvocation $MyInvocation

Write-Output "Getting [$ServiceName]"
$serviceObject = Get-WmiObject -Class Win32_Service | Where-Object {$PSItem.Name -eq $ServiceName}
If ($serviceObject)
{  
    If ($serviceObject.State -eq 'Running')
    {
        Write-Output "Stopping [$ServiceName]"
        $respone = $serviceObject.StopService()
        If ($respone.ReturnValue -ne 0)
        {
            Write-Error "Service responded with [$($respone.ReturnValue)]. See https://docs.microsoft.com/en-us/windows/desktop/cimwin32prov/stopservice-method-in-class-win32-service for details." -ErrorAction Stop
        }
    }
    $parentPath = $serviceObject.PathName | Split-Path -Parent
    Write-Output "Identified deployment location [$parentPath]"
    If ($parentPath)
    {
        If ($CleanInstall)
        {
            Write-Output "[$($MyInvocation.MyCommand.Name)]: Clean install set to [$CleanInstall], removing [$parentPath]"
            $TIMEOUT = '60'
            $cleanInstalltimer = [Diagnostics.Stopwatch]::StartNew()
            Do
            {
                Try
                {
                    Remove-Item -Path $parentPath -Force -Recurse -ErrorAction Stop
                }
                Catch
                {
                    Switch -Wildcard ($PSItem.ErrorDetails.Message)
                    {
                        '*Cannot remove*'
                        {
                            $allProcesses = Get-Process
                            $process = $allProcesses | Where-Object {$_.Path -like "$parentPath\*"} 
                            If ($process)
                            {
                                Write-Warning "[$($MyInvocation.MyCommand.Name)]: Files are still in use by [$($process.ProcessName)], killing the process!"
                                $process | Stop-Process -Force -ErrorAction SilentlyContinue
                            }
                        }
                        Default
                        {
                            Write-Error $PSItem -ErrorAction Stop
                        }
                    }
                }
                If ($cleanInstalltimer.Elapsed.TotalSeconds -gt $TIMEOUT)
                {
                    Write-Error "[$($MyInvocation.MyCommand.Name)]: [$ServiceName] did not respond within the timeout limit, clean install has failed." -ErrorAction Stop
                }
            }
            While (Test-Path $parentPath)
            $null = New-Item -ItemType Directory -Path $parentPath -Force
        }
    }
    Else
    {
        $null = New-Item -ItemType Directory -Path $parentPath -Force
    }
    Write-Output "Copying [$ArtifactPath] to [$parentPath]"
    Copy-Item -Path "$ArtifactPath\*" -Destination $parentPath -Force -Recurse
    Write-Output "Starting [$ServiceName]"
    $respone = $serviceObject.StartService()
    If ($respone.ReturnValue -ne 0)
    {
        Write-Error "Service responded with [$($respone.ReturnValue)]. See https://docs.microsoft.com/en-us/windows/desktop/cimwin32prov/startservice-method-in-class-win32-service for details." -ErrorAction Stop
    }
}
else
{
    Write-Error "Unable to locate [$ServiceName] on [$Env:ComputerName], confirm the service is installed correctly." -ErrorAction Stop   
}
Trace-VstsLeavingInvocation $MyInvocation