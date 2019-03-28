param
(
    [Parameter()]
    [string]
    $DeploymentType = (Get-VstsInput -Name 'DeploymentType' -Require),

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
    $StopProcess = (Get-VstsInput -Name 'StopProcess' -AsBool),

    [Parameter()]
    $InstallService = (Get-VstsInput -Name 'InstallService' -AsBool)
)
Trace-VstsEnteringInvocation $MyInvocation

If ($DeploymentType -eq 'Agent')
{
    $_machines = (Get-VstsInput -Name 'Machines' -Require).Split(',')
    Write-Output ("Begining deployment to [{0}]" -f $_machines -join ', ')
    $adminLogin = Get-VstsInput -Name 'AdminLogin' -Require
    $password = Get-VstsInput -Name 'Password' -Require
    $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
    $credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $adminLogin, $securePassword
}
If($InstallService)
{
    $installationPath = Get-VstsInput -Name 'InstallationPath'
    $runAsUsername = Get-VstsInput -Name 'RunAsUsername'
    $runAsPassword = Get-VstsInput -Name 'RunAsPassword'
    If($runAsPassword)
    {
        $secureRunAsPassword = ConvertTo-SecureString $runAsPassword -AsPlainText -Force
        $runAsCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $runAsUsername, $secureRunAsPassword
    }
}
$scriptBlock = {
    $serviceName      = $args[0]
    $Timeout          = $args[1]
    $StopProcess      = $args[2]
    $CleanInstall     = $args[3]
    $ArtifactPath     = $args[4]
    $installationPath = $args[5]
    $runAsCredential  = $args[6]
    Function Get-WindowsService
    {
        param
        (
            $ServiceName = $ServiceName
        )
        Get-WmiObject -Class Win32_Service | Where-Object {$PSItem.Name -eq $ServiceName}
    }
    Write-Output "Attempting to locate [$ServiceName]"
    $serviceObject = Get-WindowsService -ServiceName $ServiceName
    # If the service does not exist and the installtion path can only be provided if the Install Service flag is passed.
    If($null -eq $serviceObject -and $null -ne $installationPath)
    {
        Write-Output "Unable to locate [$ServiceName] creating a new service"
        $newService = New-Service -Name $ServiceName -BinaryPathName $installationPath -Credential $runAsCredential
        $serviceObject = Get-WindowsService -ServiceName $ServiceName
    }
    If ($serviceObject)
    {  
        If ($serviceObject.State -eq 'Running')
        {
            $stopServiceTimer = [Diagnostics.Stopwatch]::StartNew()
            Write-Output "Stopping [$ServiceName]"
            Do
            {
                $serviceObject = Get-WindowsService -ServiceName $ServiceName
                $results = $serviceObject.StopService()
                If ($stopServiceTimer.Elapsed.TotalSeconds -gt $Timeout)
                {
                    If ($StopProcess)
                    {
                        Write-Verbose "[$ServiceName] did not respond within [$Timeout] seconds, stopping process."
                        $allProcesses = Get-Process
                        $process = $allProcesses | Where-Object {$_.Path -like "$parentPath\*"}
                        If ($process)
                        {
                            Write-Warning "Files are still in use by [$($process.ProcessName)], stopping the process!"
                            $process | Stop-Process -Force -ErrorAction SilentlyContinue
                        }
                    }
                    Else
                    {
                        Write-Error "[$ServiceName] did not respond within [$Timeout] seconds." -ErrorAction Stop                    
                    }
                }
                $serviceObject = Get-WindowsService -ServiceName $ServiceName
            }
            While ($serviceObject.State -ne 'Stopped')
        }
        $parentPath = ($serviceObject.PathName | Split-Path -Parent).Replace('"', '')
        Write-Output "Identified [$ServiceName] installation directory [$parentPath]"
        If (Test-Path $parentPath)
        {
            If ($CleanInstall)
            {
                Write-Output "Clean install set to [$CleanInstall], removing [$parentPath]"
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
                                    Write-Verbose "[$ServiceName] did not respond within [$Timeout] seconds, stopping process." 
                                    $allProcesses = Get-Process
                                    $process = $allProcesses | Where-Object {$_.Path -like "$parentPath\*"} 
                                    If ($process)
                                    {
                                        Write-Warning "Files are still in use by [$($process.ProcessName)], stopping the process!"
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
                    If ($cleanInstalltimer.Elapsed.TotalSeconds -gt $Timeout)
                    {
                        Write-Error "[$ServiceName] did not respond within [$Timeout] seconds, clean install has failed." -ErrorAction Stop
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
        Copy-Item -Path "$ArtifactPath\*" -Destination $parentPath -Force -Recurse -ErrorAction Stop
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
}

$invokeCommandSplat = @{
    ScriptBlock = $scriptBlock
}
If($credential)
{
    $invokeCommandSplat.Credential = $credential
    $invokeCommandSplat.ComputerName = $_machines
}
Invoke-Command @invokeCommandSplat -ArgumentList $ServiceName, $TimeOut, $StopProcess, $CleanInstall, $ArtifactPath, $installationPath, $runAsCredential
Trace-VstsLeavingInvocation $MyInvocation