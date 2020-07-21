[Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingConvertToSecureStringWithPlainText", "")]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingUserNameAndPassWordParams", "")]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingPlainTextForPassword", "")]
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
    $RecreateService = (Get-VstsInput -Name 'RecreateService' -AsBool),

    [Parameter()]
    $StopProcess = (Get-VstsInput -Name 'StopProcess' -AsBool),

    [Parameter()]
    $InstallService = (Get-VstsInput -Name 'InstallService' -AsBool),
    
    [Parameter()]
    $StartService = (Get-VstsInput -Name 'StartService' -AsBool),

    [Parameter()]
    $Machines = (Get-VstsInput -Name 'Machines').Split(',').trim(),

    [Parameter()]
    $AdminLogin = (Get-VstsInput -Name 'AdminLogin'),

    [Parameter()]
    $Password = (Get-VstsInput -Name 'Password'),

    [Parameter()]
    $NewPsSessionOptionArguments = (Get-VstsInput -Name 'NewPsSessionOptionArguments'),

    [Parameter()]
    $UseSSL = (Get-VstsInput -Name 'UseSSL' -AsBool),

    [Parameter()]
    $InstallTopShelfService = (Get-VstsInput -Name 'InstallTopShelfService' -AsBool),

    [Parameter()]
    $InstanceName = (Get-VstsInput -Name 'InstanceName'),

    [Parameter()]
    $InstallArguments = (Get-VstsInput -Name 'InstallArguments'),

    [Parameter()]
    $StartCommand = (Get-VstsInput -Name 'InstallationPath'),

    [Parameter()]
    $ServiceDisplayName = (Get-VstsInput -Name 'ServiceDisplayName'),

    [Parameter()]
    $ServiceDescription = (Get-VstsInput -Name 'ServiceDescription'),

    [Parameter()]
    $ServiceStartupType = (Get-VstsInput -Name 'ServiceStartupType'),

    [Parameter()]
    $RunAsUsername = (Get-VstsInput -Name 'RunAsUsername'),

    [Parameter()]
    $RunAsPassword = (Get-VstsInput -Name 'RunAsPassword')


)
Trace-VstsEnteringInvocation $MyInvocation
# Import utility script
$UTILITY_PATH = Join-Path -Path $PSScriptRoot -ChildPath 'ps_modules\WindowsServiceManagerPS\WindowsServiceManagerPS.ps1'
. $UTILITY_PATH

if ($DeploymentType -eq 'Agent')
{
    Write-Output ("Begining deployment to [{0}]" -f ($Machines -join ', '))

    if (($null -ne $Password) -and ($null -ne $AdminLogin))
    {
        $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
        $credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $AdminLogin, $securePassword
    }
    $sessionOption = Get-WSMNewPSSessionOption -arguments $NewPsSessionOptionArguments
}

if ($StartCommand)
{
    $installationPath = Get-WSMFullExecuteablePath -StringContainingPath $StartCommand
}

if ($RunAsPassword)
{
    $secureRunAsPassword = ConvertTo-SecureString $RunAsPassword -AsPlainText -Force
    $runAsCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $RunAsUsername, $secureRunAsPassword
}

# Set default values of optional parameters.
if (-not $ServiceDisplayName)
{
    $ServiceDisplayName = $ServiceName
}

if (-not $ServiceDescription)
{
    $ServiceDescription = $ServiceName
}

$scriptBlock = {
    $serviceName = $args[0]
    $ServiceDisplayName = $args[1]
    $ServiceDescription = $args[2]
    $ServiceStartupType = $args[3]
    $Timeout = $args[4]
    $StopProcess = $args[5]
    $CleanInstall = $args[6]
    $ArtifactPath = $args[7]
    $InstallationPath = $args[8]
    $StartCommand = $args[9]
    $runAsCredential = $args[10]
    $InstallTopShelfService = $args[11]
    $InstanceName = $args[12]
    $InstallArguments = $args[13]
    $startService = $args[14]
    $RecreateService = $args[15]
    $InstallService = $args[16]
    
    # Handle TopShelf service instance name
    if ($InstanceName.Length -ne 0)
    {
        Write-Output "[$env:ComputerName]: Instance Name: [$InstanceName]"
        $serviceName = "{0}`${1}" -f $ServiceName.split('$')[0], $InstanceName
    }
    
    Write-Output "[$env:ComputerName]: Attempting to locate [$ServiceName]"
    $serviceObject = Get-WSMWindowsService -ServiceName $ServiceName
    if ($serviceObject)
    {
        if($RecreateService)
        {
            Write-Output "[$env:ComputerName]: Recreate service set to [$RecreateService]"
            Write-Output "[$env:ComputerName]: Stopping the service [$ServiceName]"
            $serviceObject = Stop-WSMWindowsService -ServiceName $ServiceName
            Write-Output "[$env:ComputerName]: Deleting the service [$ServiceName]"
            
                $deleteResults = $serviceObject.Delete()
            if($deleteResults.ReturnValue -eq 0)
            {
                Write-Output "[$env:ComputerName]: Successfully removed [$ServiceName]"  
            }
            else 
            {
                return Write-Error "[$env:ComputerName]: Unable to remove [$ServiceName] returned value was [$($deleteResults.ReturnValue)]. See https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/delete-method-in-class-win32-service for more information on return values" 
            }
        }
    }
    else
    {
        Write-Output "[$env:ComputerName]: Unable to locate [$ServiceName]."
    }

    # Install service
    if((-not $serviceObject) -and $InstallService)
    {
        $parentPath = Get-WSMFullExecuteablePath -StringContainingPath $InstallationPath -JustParentPath
        New-WSMServiceDirectory -ParentPath $parentPath
        Write-Output "[$env:ComputerName]: Creating service [$ServiceName]."
        if ($InstallTopShelfService)
        {
            # Binaires needed to complete install for TopShelf
            Copy-WSMServiceBinaries -ArtifactPath $ArtifactPath -ParentPath $parentPath -ErrorAction Stop
            $arguments = @(
                'install'
                '-servicename:{0}' -f $ServiceName.split('$')[0]
            )
            if ($runAsCredential)
            {
                $arguments += '-username:{0}' -f $runAsCredential.UserName
                $arguments += '-password:{0}' -f $runAsCredential.GetNetworkCredential().Password
            }
            if ($InstanceName)
            {
                $arguments += '-instance:{0}' -f $InstanceName
            }
            if ($InstallArguments)
            {
                $arguments += $InstallArguments
            }

            Write-Output "[$env:ComputerName]: Installing topshelf with arguments $arguments"
            & $installationPath $arguments
            $freshTopShelfInstall = $true
        }
        else
        {
            if ($ServiceStartupType -eq "Delayed")
            {
                $startupType = "Automatic"
                $delayed = $true
            }
            else
            {
                $startupType = $ServiceStartupType
                $delayed = $false
            }
            $newServiceSplat = @{
                Name           = $ServiceName
                BinaryPathName = $StartCommand
                DisplayName    = $ServiceDisplayName
                StartupType    = $startupType
                Description    = $ServiceDescription
            }
            if ($runAsCredential)
            {
                Write-Output "[$env:ComputerName]: Setting RunAsCredentials"
                $newServiceSplat.Credential = $runAsCredential
                Write-Output "[$env:ComputerName]: Granting [$($runAsCredential.UserName)] logon as a service rights"
                Add-LocalUserToLogonAsAService -user $runAsCredential.UserName
            }
            $null = New-Service @newServiceSplat            
            Write-Output "[$env:ComputerName]: Service [$ServiceName] created."

            if ($delayed)
            {
                Write-Output "[$env:ComputerName]: Set [$ServiceName] to delayed start"
                Start-Process -FilePath sc.exe -ArgumentList "config ""$ServiceName"" start=delayed-auto"
            }
        }
    }     

    $serviceObject = Get-WSMWindowsService -ServiceName $ServiceName
    if ($freshTopShelfInstall)
    {
        # Topshelf installation completed the file copy so skip the clean install process
        if ($startService)
        {
            Start-WSMWindowsService -ServiceName $ServiceName
        }
    }
    elseif ($serviceObject)
    {  
        $serviceObject = Stop-WSMWindowsService -ServiceName $ServiceName -Timeout $Timeout -StopProcess:$StopProcess
        $parentPath = Get-WSMFullExecuteablePath -StringContainingPath $serviceObject.PathName -JustParentPath
        Write-Output "[$env:ComputerName]: Identified [$ServiceName] installation directory [$parentPath]"

        if (Test-Path $parentPath)
        {
            
            if ($CleanInstall)
            {
                Write-Output "[$env:ComputerName]: Clean install set to [$CleanInstall], removing the contents of [$parentPath]"
                Invoke-WSMCleanInstall -ParentPath $parentPath -StopProcess $StopProcess -Timeout $TimeOut -ServiceName $ServiceName
            }
        }
        else
        {
            New-WSMServiceDirectory -ParentPath $parentPath
        }
        Copy-WSMServiceBinaries -ArtifactPath $ArtifactPath -ParentPath $parentPath -ErrorAction Stop
        
        if ($startService)
        {
            Start-WSMWindowsService -ServiceName $ServiceName
        }
    }
    else
    {
        return Write-Error "[$env:ComputerName]: Unable to locate [$ServiceName], confirm the service is installed correctly." 
    }
}


if ($Machines)
{
    $newPSSessionSPlat = @{
        ComputerName = $Machines 
        SessionOption = $sessionOption 
        UseSSL = $UseSSL
    }
    if ($credential)
    {
        $newPSSessionSPlat.Credential = $credential
    }
    $sessions = New-PSSession @newPSSessionSPlat
}

$invokeCommandSplat = @{
    Session = $sessions
    ArgumentList = $ServiceName, 
        $ServiceDisplayName, 
        $ServiceDescription, 
        $ServiceStartupType, 
        $TimeOut, 
        $StopProcess, 
        $CleanInstall,
        $ArtifactPath,
        $installationPath,
        $StartCommand,
        $runAsCredential,
        $InstallTopShelfService,
        $InstanceName,
        $InstallArguments,
        $startService,
        $RecreateService,
        $InstallService
}

    # Import utility script into session
    Invoke-Command @invokeCommandSplat -FilePath $UTILITY_PATH

# Invoke script block
Invoke-Command @invokeCommandSplat -ScriptBlock $scriptBlock
Trace-VstsLeavingInvocation $MyInvocation