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
    $InstallService = (Get-VstsInput -Name 'InstallService' -AsBool),
    
    [Parameter()]
    $StartService = (Get-VstsInput -Name 'StartService' -AsBool)
)
Trace-VstsEnteringInvocation $MyInvocation

. "$PSScriptRoot\Utility.ps1"

If ($DeploymentType -eq 'Agent')
{
    $_machines = (Get-VstsInput -Name 'Machines' -Require).Split(',').trim()
    Write-Output ("Begining deployment to [{0}]" -f ($_machines -join ', '))
    $adminLogin = (Get-VstsInput -Name 'AdminLogin' -Require )
    $password = (Get-VstsInput -Name 'Password' -Require )
    $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
    $credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $adminLogin, $securePassword
    $input_NewPsSessionOptionArguments = (Get-VstsInput -Name "NewPsSessionOptionArguments")
    $sessionOption = Get-NewPSSessionOption -arguments $input_NewPsSessionOptionArguments
    $useSSL = (Get-VstsInput -Name 'UseSSL' -AsBool)
}
If($InstallService)
{
    $installTopShelfService = (Get-VstsInput -Name 'InstallTopShelfService' -AsBool )
    If($installTopShelfService)
    {
        $instanceName = (Get-VstsInput -Name 'InstanceName' )
        $installArguments = (Get-VstsInput -Name 'InstallArguments' )
    }
    $installationPath = (Get-VstsInput -Name 'InstallationPath' )
    If(-not($installationPath.EndsWith('.exe')))
    {
        return Write-Error -Message "The installation path parameter should end with an '.exe'. InstallationPath should be populated with a path to the service executable but it is currently [$InstallationPath]."
    }
    $runAsUsername = (Get-VstsInput -Name 'RunAsUsername' )
    $runAsPassword = (Get-VstsInput -Name 'RunAsPassword' )

    If($runAsPassword)
    {
        $secureRunAsPassword = ConvertTo-SecureString $runAsPassword -AsPlainText -Force
        $runAsCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $runAsUsername, $secureRunAsPassword
    }
}
$scriptBlock = {
    $serviceName            = $args[0]
    $Timeout                = $args[1]
    $StopProcess            = $args[2]
    $CleanInstall           = $args[3]
    $ArtifactPath           = $args[4]
    $installationPath       = $args[5]
    $runAsCredential        = $args[6]
    $installTopShelfService = $args[7]
    $instanceName           = $args[8]
    $installArguments       = $args[9]
    If($instanceName.Length -ne 0)
    {
        Write-Output "[$env:ComputerName]: Instance Name: [$instanceName]"
        $serviceName = "{0}`${1}" -f $ServiceName.split('$')[0], $instanceName
    }
    Function Get-WindowsService
    {
        param
        (
            $ServiceName
        )
        Get-WmiObject -Class Win32_Service | Where-Object {$PSItem.Name -eq $ServiceName}
    }
    Function Start-WindowsService
    {
        param
        (
            $ServiceName
        )
        Write-Output "[$env:ComputerName]: Starting [$ServiceName]"
        $serviceObject = Get-WindowsService -ServiceName $ServiceName
        $respone = $serviceObject.StartService()
         If ($respone.ReturnValue -ne 0)
         {
             return Write-Error -Message "[$env:ComputerName]: Service responded with [$($respone.ReturnValue)]. See https://docs.microsoft.com/en-us/windows/desktop/cimwin32prov/startservice-method-in-class-win32-service for details."
         }
         else 
         {
             Write-Output "[$env:ComputerName]: [$ServiceName] started successfully!"
         }
    }
    Write-Output "[$env:ComputerName]: Attempting to locate [$ServiceName]"
    $serviceObject = Get-WindowsService -ServiceName $ServiceName
    # If the service does not exist and the installtion path can only be provided if the Install Service flag is passed.
    If($null -eq $serviceObject -and $null -ne $installationPath)
    {
        Write-Output "[$env:ComputerName]: Unable to locate [$ServiceName] creating a new service"
        If($installTopShelfService)
        {
            $parentPath = $installationPath | Split-Path -Parent
            If(-not(Test-Path $parentPath))
            {
                $null = New-Item -Path $parentPath -ItemType 'Directory' -Force
            }
            Write-Output "[$env:ComputerName]: Copying [$ArtifactPath] to [$parentPath]"
            Copy-Item -Path "$ArtifactPath\*" -Destination $parentPath -Force -Recurse -ErrorAction Stop
            $arguments = @(
                'install'
                '-servicename:{0}' -f $ServiceName.split('$')[0]
            )
            If($runAsCredential)
            {
                $arguments += '-username:{0}' -f $runAsCredential.UserName
                $arguments += '-password:{0}' -f $runAsCredential.GetNetworkCredential().Password
            }
            If($instanceName)
            {
                $arguments += '-instance:{0}' -f $instanceName
            }
            If($installArguments)
            {
                $arguments += $installArguments
            }
            Write-Host "[$env:ComputerName]: Installing topshelf with arguments $arguments"
            & $installationPath $arguments
            $freshTopShelfInstall = $true
        }
        Else
        {
            $newServiceSplat = @{
                Name = $ServiceName
                BinaryPathName = $installationPath
            }
            If($runAsCredential)
            {
                $newServiceSplat.Credential = $runAsCredential
            }
            $newService = New-Service @newServiceSplat
        }
    }
    $serviceObject = Get-WindowsService -ServiceName $ServiceName
    If($freshTopShelfInstall)
    {
        # Topshelf installation completed the file copy so skip the clean install process
        
        If ($StartService)
        {
            Start-WindowsService -ServiceName $ServiceName
        }
    }
    ElseIf ($serviceObject)
    {  
        If ($serviceObject.State -eq 'Running')
        {
            $stopServiceTimer = [Diagnostics.Stopwatch]::StartNew()
            Write-Output "[$env:ComputerName]: Stopping [$ServiceName]"
            Do
            {
                $serviceObject = Get-WindowsService -ServiceName $ServiceName
                $results = $serviceObject.StopService()
                If ($stopServiceTimer.Elapsed.TotalSeconds -gt $Timeout)
                {
                    If ($StopProcess)
                    {
                        Write-Verbose "[$env:ComputerName]: [$ServiceName] did not respond within [$Timeout] seconds, stopping process."
                        $allProcesses = Get-Process
                        $process = $allProcesses | Where-Object {$_.Path -like "$parentPath\*"}
                        If ($process)
                        {
                            Write-Warning -Message "[$env:ComputerName]: Files are still in use by [$($process.ProcessName)], stopping the process!"
                            $process | Stop-Process -Force -ErrorAction SilentlyContinue
                        }
                    }
                    Else
                    {
                        return Write-Error -Message "[$env:ComputerName]: [$ServiceName] did not respond within [$Timeout] seconds."             
                    }
                }
                $serviceObject = Get-WindowsService -ServiceName $ServiceName
            }
            While ($serviceObject.State -ne 'Stopped')
        }
        $parentPath = ($serviceObject.PathName | Split-Path -Parent).Replace('"', '')
        Write-Output "[$env:ComputerName]: Identified [$ServiceName] installation directory [$parentPath]"
        If (Test-Path $parentPath)
        {
            If ($CleanInstall)
            {
                Write-Output "[$env:ComputerName]: Clean install set to [$CleanInstall], removing the contents of [$parentPath]"
                $cleanInstalltimer = [Diagnostics.Stopwatch]::StartNew()
                Do
                {
                    Try
                    {
                        Get-ChildItem -Path $parentPath -Recurse -Force | Remove-Item -Recurse -Force -ErrorAction Stop
                    }
                    Catch
                    {
                        Switch -Wildcard ($PSItem.ErrorDetails.Message)
                        {
                            '*Cannot remove*'
                            {
                                If ($StopProcess)
                                {
                                    Write-Verbose "[$env:ComputerName]: [$ServiceName] did not respond within [$Timeout] seconds, stopping process." 
                                    $allProcesses = Get-Process
                                    $process = $allProcesses | Where-Object {$_.Path -like "$parentPath\*"} 
                                    If ($process)
                                    {
                                        Write-Warning -Message "[$env:ComputerName]: Files are still in use by [$($process.ProcessName)], stopping the process!"
                                        $process | Stop-Process -Force -ErrorAction SilentlyContinue
                                    }
                                }
                                else
                                {
                                    return Write-Error -Message $PSItem
                                }
    
                            }
                            Default
                            {
                                return Write-Error -Message $PSItem
                            }
                        }
                    }
                    If ($cleanInstalltimer.Elapsed.TotalSeconds -gt $Timeout)
                    {
                        return Write-Error -Message "[$env:ComputerName]: [$ServiceName] did not respond within [$Timeout] seconds, clean install has failed."
                    }
                }
                While (Get-ChildItem -Path $parentPath -Recurse -Force)
                $null = New-Item -ItemType Directory -Path $parentPath -Force
            }
        }
        Else
        {
            $null = New-Item -ItemType Directory -Path $parentPath -Force
        }
        Write-Output "[$env:ComputerName]: Copying [$ArtifactPath] to [$parentPath]"
        Copy-Item -Path "$ArtifactPath\*" -Destination $parentPath -Force -Recurse -ErrorAction Stop
        
        If($StartService)
        {
            Start-WindowsService -ServiceName $ServiceName
        }

    }
    else
    {
        return Write-Error "[$env:ComputerName]: Unable to locate [$ServiceName], confirm the service is installed correctly." 
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
if($sessionOption)
{
    $invokeCommandSplat.sessionOption = $sessionOption
}
if($useSSL)
{
    $invokeCommandSplat.UseSSL = $true
}
Invoke-Command @invokeCommandSplat -ArgumentList $ServiceName, $TimeOut, $StopProcess, $CleanInstall, $ArtifactPath, $installationPath, $runAsCredential, $installTopShelfService, $instanceName, $installArguments
Trace-VstsLeavingInvocation $MyInvocation