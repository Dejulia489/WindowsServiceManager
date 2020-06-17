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

if ($DeploymentType -eq 'Agent') {
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

if ($InstallService) {
    $installTopShelfService = (Get-VstsInput -Name 'InstallTopShelfService' -AsBool )
    if ($installTopShelfService) {
        $instanceName = (Get-VstsInput -Name 'InstanceName' )
        $installArguments = (Get-VstsInput -Name 'InstallArguments' )
    }
    $startCommand = (Get-VstsInput -Name 'InstallationPath' )

    $serviceDisplayName = (Get-VstsInput -Name 'ServiceDisplayName')
    $serviceDescription = (Get-VstsInput -Name 'ServiceDescription')
    $serviceStartupType = (Get-VstsInput -Name 'ServiceStartupType')

    $installationPath = Get-FullExecuteablePath -StringContainingPath $startCommand

    $runAsUsername = (Get-VstsInput -Name 'RunAsUsername' )
    $runAsPassword = (Get-VstsInput -Name 'RunAsPassword' )

    if ($runAsPassword) {
        $secureRunAsPassword = ConvertTo-SecureString $runAsPassword -AsPlainText -Force
        $runAsCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $runAsUsername, $secureRunAsPassword
    }
}

# Set default values of optional parameters.
if (-not $ServiceDisplayName) {
    $ServiceDisplayName = $ServiceName
}

# fix ServiceName (can not contain Spaces)
if ($ServiceName.Contains(' ')) {
    $ServiceName = $ServiceName.Replace(' ', '')
}

$scriptBlock = {
    $serviceName =              $args[0]
    $serviceDisplayName =       $args[1]
    $serviceDescription =       $args[2]
    $serviceStartupType =       $args[3]
    $Timeout =                  $args[4]
    $StopProcess =              $args[5]
    $CleanInstall =             $args[6]
    $ArtifactPath =             $args[7]
    $installationPath =         $args[8]
    $startCommand =             $args[9]
    $runAsCredential =          $args[10]
    $installTopShelfService =   $args[11]
    $instanceName =             $args[12]
    $installArguments =         $args[13]
    $startService =             $args[14]

    if ($instanceName.Length -ne 0) {
        Write-Output "[$env:ComputerName]: Instance Name: [$instanceName]"
        $serviceName = "{0}`${1}" -f $ServiceName.split('$')[0], $instanceName
    }
    
    Write-Output "[$env:ComputerName]: Attempting to locate [$ServiceName]"
    $serviceObject = Get-WindowsService -ServiceName $ServiceName
    # If the service does not exist or cleanInstall is enabled and the installtion path can only be provided if the Install Service flag is passed.
    if (($null -eq $serviceObject -or $CleanInstall) -and $null -ne $installationPath) {
        if ($serviceObject) {
            Write-Output "[$env:ComputerName]: Clean install set to [$CleanInstall], removing the Service [$ServiceName]"

            $serviceObject = Stop-WindowsService -ServiceName $ServiceName

            $serviceObject.Delete()
            Write-Output "[$env:ComputerName]: Removed Service [$ServiceName]"            
        }
        else {
            Write-Output "[$env:ComputerName]: Unable to locate [$ServiceName] creating a new service"
        }

        if ($installTopShelfService) {
            $parentPath = Get-FullExecuteablePath -StringContainingPath $installationPath -JustParentPath
            if (-not(Test-Path $parentPath)) {
                $null = New-Item -Path $parentPath -ItemType 'Directory' -Force
            }

            Write-Output "[$env:ComputerName]: Copying [$ArtifactPath] to [$parentPath]"
            Copy-Item -Path "$ArtifactPath\*" -Destination $parentPath -Force -Recurse -ErrorAction Stop

            $arguments = @(
                'install'
                '-servicename:{0}' -f $ServiceName.split('$')[0]
            )
            if ($runAsCredential) {
                $arguments += '-username:{0}' -f $runAsCredential.UserName
                $arguments += '-password:{0}' -f $runAsCredential.GetNetworkCredential().Password
            }
            if ($instanceName) {
                $arguments += '-instance:{0}' -f $instanceName
            }
            if ($installArguments) {
                $arguments += $installArguments
            }

            Write-Host "[$env:ComputerName]: Installing topshelf with arguments $arguments"
            & $installationPath $arguments
            $freshTopShelfInstall = $true
        }
        else {
            Write-Output "[$env:ComputerName]: Start creating Service [$ServiceName]."
            if ($serviceStartupType -eq "Delayed") {
                $startupType = "Automatic"
                $delayed = $true
            }
            else {
                $startupType = $serviceStartupType
                $delayed = $false
            }

            $newServiceSplat = @{
                Name           = $ServiceName
                BinaryPathName = $startCommand
                DisplayName    = $serviceDisplayName
                StartupType    = $startupType
            }

            # add Description just if Descripion is provided to prevent Parameter null or empty Exception
            if($serviceDescription) {                
                Write-Output "[$env:ComputerName]: Adding Description [$serviceDescription]"
                $newServiceSplat.Description = $serviceDescription
            }

            if ($runAsCredential) {
                Write-Output "[$env:ComputerName]: Setting RunAsCredentials"
                $newServiceSplat.Credential = $runAsCredential
                # load Function
                . "$PSScriptRoot\Add-LocalUserToLogonAsAService.ps1"
                Add-LocalUserToLogonAsAService -user $runAsCredential.UserName
            }
            $newService = New-Service @newServiceSplat            
            Write-Output "[$env:ComputerName]: Service [$ServiceName] created."

            if ($delayed) {
                Write-Output "[$env:ComputerName]: Set [$ServiceName] to Delayed start"
                Start-Process -FilePath sc.exe -ArgumentList "config ""$ServiceName"" start=delayed-auto"
            }
        }
    }

    $serviceObject = Get-WindowsService -ServiceName $ServiceName
    
    if ($freshTopShelfInstall) {
        # Topshelf installation completed the file copy so skip the clean install process
        
        if ($startService) {
            Start-WindowsService -ServiceName $ServiceName
        }
    }
    elseif ($serviceObject) {  
        $serviceObject = Stop-WindowsService -ServiceName $ServiceName -Timeout $Timeout -StopProcess:$StopProcess
        $parentPath = Get-FullExecuteablePath -StringContainingPath $serviceObject.PathName -JustParentPath
        Write-Output "[$env:ComputerName]: Identified [$ServiceName] installation directory [$parentPath]"

        if (Test-Path $parentPath) {
            if ($CleanInstall) {
                Write-Output "[$env:ComputerName]: Clean install set to [$CleanInstall], removing the contents of [$parentPath]"
                $cleanInstalltimer = [Diagnostics.Stopwatch]::StartNew()
                do {
                    try {
                        Get-ChildItem -Path $parentPath -Recurse -Force | Remove-Item -Recurse -Force -ErrorAction Stop
                    }
                    catch {
                        switch -Wildcard ($PSItem.ErrorDetails.Message) {
                            '*Cannot remove*' {
                                if ($StopProcess) {
                                    Write-Verbose "[$env:ComputerName]: [$ServiceName] did not respond within [$Timeout] seconds, stopping process." 
                                    $allProcesses = Get-Process
                                    $process = $allProcesses | Where-Object { $_.Path -like "$parentPath\*" } 
                                    if ($process) {
                                        Write-Warning -Message "[$env:ComputerName]: Files are still in use by [$($process.ProcessName)], stopping the process!"
                                        $process | Stop-Process -Force -ErrorAction SilentlyContinue
                                    }
                                }
                                else {
                                    return Write-Error -Message $PSItem
                                }    
                            }
                            default {
                                return Write-Error -Message $PSItem
                            }
                        }
                    }
                    if ($cleanInstalltimer.Elapsed.TotalSeconds -gt $Timeout) {
                        return Write-Error -Message "[$env:ComputerName]: [$ServiceName] did not respond within [$Timeout] seconds, clean install has failed."
                    }
                }
                while (Get-ChildItem -Path $parentPath -Recurse -Force)
                $null = New-Item -ItemType Directory -Path $parentPath -Force
            }
        }
        else {
            $null = New-Item -ItemType Directory -Path $parentPath -Force
        }

        Write-Output "[$env:ComputerName]: Copying [$ArtifactPath] to [$parentPath]"
        Copy-Item -Path "$ArtifactPath\*" -Destination $parentPath -Force -Recurse -ErrorAction Stop
        
        if ($startService) {
            Start-WindowsService -ServiceName $ServiceName
        }
    }
    else {
        return Write-Error "[$env:ComputerName]: Unable to locate [$ServiceName], confirm the service is installed correctly." 
    }
}

$invokeCommandSplat = @{
    ScriptBlock = $scriptBlock
}
if ($credential) {
    $invokeCommandSplat.Credential = $credential
    $invokeCommandSplat.ComputerName = $_machines
}
if ($sessionOption) {
    $invokeCommandSplat.sessionOption = $sessionOption
}
if ($useSSL) {
    $invokeCommandSplat.UseSSL = $true
}

Invoke-Command @invokeCommandSplat -ArgumentList $ServiceName, $serviceDisplayName, $serviceDescription, $serviceStartupType, $TimeOut, $StopProcess, $CleanInstall, $ArtifactPath, $installationPath, $startCommand, $runAsCredential, $installTopShelfService, $instanceName, $installArguments, $startService
Trace-VstsLeavingInvocation $MyInvocation