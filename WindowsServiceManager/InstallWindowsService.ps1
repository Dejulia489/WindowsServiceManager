param
(
    [Parameter()]
    [string]
    $ServiceName = (Get-VstsInput -Name 'ServiceName' -Require), 

    [Parameter()]
    [string]
    $ArtifactPath = (Get-VstsInput -Name 'ArtifactPath' -Require),

    [Parameter()]
    [string]
    $TimeOut = (Get-VstsInput -Name 'TimeOut' -Require),

    [Parameter()]
    $CleanInstall = (Get-VstsInput -Name 'CleanInstall' -AsBool),

    [Parameter()]
    $StopProcess = (Get-VstsInput -Name 'StopProcess' -AsBool)
)
Trace-VstsEnteringInvocation $MyInvocation

Write-Output "Getting [$ServiceName]"
$serviceObject = Get-WmiObject -Class Win32_Service | Where-Object {$PSItem.Name -eq $ServiceName}
If ($serviceObject)
{  
    If ($serviceObject.State -eq 'Running')
    {
        $stopServiceTimer = [Diagnostics.Stopwatch]::StartNew()
        Write-Output "Stopping [$ServiceName]"
        Do
        {
            $serviceObject = Get-WmiObject -Class Win32_Service | Where-Object {$PSItem.Name -eq $ServiceName}
            $results = $serviceObject.StopService()
            If ($stopServiceTimer.Elapsed.TotalSeconds -gt $TimeOut)
            {
                If ($StopProcess)
                {
                    Write-Verbose "[$ServiceName] did not respond within the timeout limit, stopping process."
                    $allProcesses = Get-Process
                    $process = $allProcesses | Where-Object {$_.Path -like "$parentPath\*"}
                    If ($process)
                    {
                        Write-Output "Killing [$($process.ProcessName)]!"
                        $process | Stop-Process -Force -ErrorAction SilentlyContinue
                    }
                }
                Else
                {
                    Write-Error "[$($MyInvocation.MyCommand.Name)]: [$ServiceName] did not respond within the timeout limit." -ErrorAction Stop                    
                }
            }
            $serviceObject = Get-WmiObject -Class Win32_Service | Where-Object {$PSItem.Name -eq $ServiceName}
        }
        While ($serviceObject.State -ne 'Stopped')
    }
    $parentPath = ($serviceObject.PathName | Split-Path -Parent).Replace('"', '')
    Write-Output "Identified deployment location [$parentPath]"
    If (Test-Path $parentPath)
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
                            If ($StopProcess)
                            {
                                $allProcesses = Get-Process
                                $process = $allProcesses | Where-Object {$_.Path -like "$parentPath\*"} 
                                If ($process)
                                {
                                    Write-Warning "[$($MyInvocation.MyCommand.Name)]: Files are still in use by [$($process.ProcessName)], stopping the process!"
                                    $process | Stop-Process -Force -ErrorAction SilentlyContinue
                                }
                            }
                            else
                            {
                                Write-Error $PSItem -ErrorAction Stop
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