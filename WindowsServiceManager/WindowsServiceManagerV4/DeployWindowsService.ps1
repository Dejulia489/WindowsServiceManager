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
    $StartService = (Get-VstsInput -Name 'StartService' -AsBool)
)
Trace-VstsEnteringInvocation $MyInvocation

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
    Write-Host "[$env:ComputerName]: Starting [$ServiceName]"
    $serviceObject = Get-WindowsService -ServiceName $ServiceName
    $respone = $serviceObject.StartService()
    if ($respone.ReturnValue -ne 0) {
        return Write-Error -Message "[$env:ComputerName]: Service responded with [$($respone.ReturnValue)]. See https://docs.microsoft.com/en-us/windows/desktop/cimwin32prov/startservice-method-in-class-win32-service for details."
    }
    else {
        Write-Host "[$env:ComputerName]: [$ServiceName] started successfully!"
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
        Write-Host "[$env:ComputerName]: Stopping Service [$ServiceName]"
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

        Write-Host "[$env:ComputerName]: Stopped Service [$ServiceName]"
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
    $matchPattern = '( |^)(?<path>([a-zA-Z]):\\([\\\w\/\(\)\[\]{}öäüÖÄÜ°^!§$%&=`´,;@#+._-]+)(.exe|.dll))|(( "|^")(?<path2>(([a-zA-Z]):\\([\\\w\/\(\)\[\]{}öäüÖÄÜ°^!§$%&=`´,;@#+._ -]+)(.exe|.dll)))(" |"$))'

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


if ($DeploymentType -eq 'Agent')
{
    $_machines = (Get-VstsInput -Name 'Machines' -Require).Split(',').trim()
    Write-Host ("Begining deployment to [{0}]" -f ($_machines -join ', '))
    $adminLogin = (Get-VstsInput -Name 'AdminLogin' -Require )
    $password = (Get-VstsInput -Name 'Password' -Require )
    $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
    $credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $adminLogin, $securePassword
    $input_NewPsSessionOptionArguments = (Get-VstsInput -Name "NewPsSessionOptionArguments")
    $sessionOption = Get-NewPSSessionOption -arguments $input_NewPsSessionOptionArguments
    $useSSL = (Get-VstsInput -Name 'UseSSL' -AsBool)
}

if ($InstallService)
{
    $installTopShelfService = (Get-VstsInput -Name 'InstallTopShelfService' -AsBool )
    if ($installTopShelfService)
    {
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

    if ($runAsPassword)
    {
        $secureRunAsPassword = ConvertTo-SecureString $runAsPassword -AsPlainText -Force
        $runAsCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $runAsUsername, $secureRunAsPassword
    }
}

# Set default values of optional parameters.
if (-not $ServiceDisplayName)
{
    $ServiceDisplayName = $ServiceName
}

# fix ServiceName (can not contain Spaces)
if ($ServiceName.Contains(' '))
{
    $ServiceName = $ServiceName.Replace(' ', '')
}

$scriptBlock = {
    $serviceName = $args[0]
    $serviceDisplayName = $args[1]
    $serviceDescription = $args[2]
    $serviceStartupType = $args[3]
    $Timeout = $args[4]
    $StopProcess = $args[5]
    $CleanInstall = $args[6]
    $ArtifactPath = $args[7]
    $installationPath = $args[8]
    $startCommand = $args[9]
    $runAsCredential = $args[10]
    $installTopShelfService = $args[11]
    $instanceName = $args[12]
    $installArguments = $args[13]
    $startService = $args[14]
    $RecreateService = $args[15]

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
        Write-Host "[$env:ComputerName]: Starting [$ServiceName]"
        $serviceObject = Get-WindowsService -ServiceName $ServiceName
        $respone = $serviceObject.StartService()
        if ($respone.ReturnValue -ne 0) {
            return Write-Error -Message "[$env:ComputerName]: Service responded with [$($respone.ReturnValue)]. See https://docs.microsoft.com/en-us/windows/desktop/cimwin32prov/startservice-method-in-class-win32-service for details."
        }
        else {
            Write-Host "[$env:ComputerName]: [$ServiceName] started successfully!"
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
            Write-Host "[$env:ComputerName]: Stopping Service [$ServiceName]"
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
    
            Write-Host "[$env:ComputerName]: Stopped Service [$ServiceName]"
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
        $matchPattern = '( |^)(?<path>([a-zA-Z]):\\([\\\w\/\(\)\[\]{}öäüÖÄÜ°^!§$%&=`´,;@#+._-]+)(.exe|.dll))|(( "|^")(?<path2>(([a-zA-Z]):\\([\\\w\/\(\)\[\]{}öäüÖÄÜ°^!§$%&=`´,;@#+._ -]+)(.exe|.dll)))(" |"$))'
    
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
    
    if ($instanceName.Length -ne 0)
    {
        Write-Host "[$env:ComputerName]: Instance Name: [$instanceName]"
        $serviceName = "{0}`${1}" -f $ServiceName.split('$')[0], $instanceName
    }
    
    Write-Host "[$env:ComputerName]: Attempting to locate [$ServiceName]"
    $serviceObject = Get-WindowsService -ServiceName $ServiceName
    # If the service does not exist or cleanInstall is enabled and the installtion path can only be provided if the Install Service flag is passed.
    if (($null -eq $serviceObject -or $RecreateService) -and $null -ne $installationPath)
    {
        if ($serviceObject)
        {
            Write-Host "[$env:ComputerName]: Recreate service set to [$RecreateService], removing the Service [$ServiceName]"
            $serviceObject = Stop-WindowsService -ServiceName $ServiceName
            $serviceObject.Delete()
            Write-Host "[$env:ComputerName]: Removed Service [$ServiceName]"  
        }
        else
        {
            Write-Host "[$env:ComputerName]: Unable to locate [$ServiceName] creating a new service"
        }
        if ($installTopShelfService)
        {
            $parentPath = Get-FullExecuteablePath -StringContainingPath $installationPath -JustParentPath
            if (-not(Test-Path $parentPath))
            {
                $null = New-Item -Path $parentPath -ItemType 'Directory' -Force
            }

            Write-Host "[$env:ComputerName]: Copying [$ArtifactPath] to [$parentPath]"
            Copy-Item -Path "$ArtifactPath\*" -Destination $parentPath -Force -Recurse -ErrorAction Stop

            $arguments = @(
                'install'
                '-servicename:{0}' -f $ServiceName.split('$')[0]
            )
            if ($runAsCredential)
            {
                $arguments += '-username:{0}' -f $runAsCredential.UserName
                $arguments += '-password:{0}' -f $runAsCredential.GetNetworkCredential().Password
            }
            if ($instanceName)
            {
                $arguments += '-instance:{0}' -f $instanceName
            }
            if ($installArguments)
            {
                $arguments += $installArguments
            }

            Write-Host "[$env:ComputerName]: Installing topshelf with arguments $arguments"
            & $installationPath $arguments
            $freshTopShelfInstall = $true
        }
        else
        {
            Write-Host "[$env:ComputerName]: Start creating Service [$ServiceName]."
            if ($serviceStartupType -eq "Delayed")
            {
                $startupType = "Automatic"
                $delayed = $true
            }
            else
            {
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
            if ($serviceDescription)
            {                
                Write-Host "[$env:ComputerName]: Adding Description [$serviceDescription]"
                $newServiceSplat.Description = $serviceDescription
            }

            if ($runAsCredential)
            {
                Write-Host "[$env:ComputerName]: Setting RunAsCredentials"
                $newServiceSplat.Credential = $runAsCredential
                # load Function
                . "$PSScriptRoot\Add-LocalUserToLogonAsAService.ps1"
                Add-LocalUserToLogonAsAService -user $runAsCredential.UserName
            }
            $newService = New-Service @newServiceSplat            
            Write-Host "[$env:ComputerName]: Service [$ServiceName] created."

            if ($delayed)
            {
                Write-Host "[$env:ComputerName]: Set [$ServiceName] to Delayed start"
                Start-Process -FilePath sc.exe -ArgumentList "config ""$ServiceName"" start=delayed-auto"
            }
        }
    }

    $serviceObject = Get-WindowsService -ServiceName $ServiceName
    
    if ($freshTopShelfInstall)
    {
        # Topshelf installation completed the file copy so skip the clean install process
        
        if ($startService)
        {
            Start-WindowsService -ServiceName $ServiceName
        }
    }
    elseif ($serviceObject)
    {  
        $serviceObject = Stop-WindowsService -ServiceName $ServiceName -Timeout $Timeout -StopProcess:$StopProcess
        $parentPath = Get-FullExecuteablePath -StringContainingPath $serviceObject.PathName -JustParentPath
        Write-Host "[$env:ComputerName]: Identified [$ServiceName] installation directory [$parentPath]"

        if (Test-Path $parentPath)
        {
            if ($CleanInstall)
            {
                Write-Host "[$env:ComputerName]: Clean install set to [$CleanInstall], removing the contents of [$parentPath]"
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
                                        Write-Warning -Message "[$env:ComputerName]: Files are still in use by [$($process.ProcessName)], stopping the process!"
                                        $process | Stop-Process -Force -ErrorAction SilentlyContinue
                                    }
                                }
                                else
                                {
                                    return Write-Error -Message $PSItem
                                }    
                            }
                            default
                            {
                                return Write-Error -Message $PSItem
                            }
                        }
                    }
                    if ($cleanInstalltimer.Elapsed.TotalSeconds -gt $Timeout)
                    {
                        return Write-Error -Message "[$env:ComputerName]: [$ServiceName] did not respond within [$Timeout] seconds, clean install has failed."
                    }
                }
                while (Get-ChildItem -Path $parentPath -Recurse -Force)
                $null = New-Item -ItemType Directory -Path $parentPath -Force
            }
        }
        else
        {
            $null = New-Item -ItemType Directory -Path $parentPath -Force
        }

        Write-Host "[$env:ComputerName]: Copying [$ArtifactPath] to [$parentPath]"
        Copy-Item -Path "$ArtifactPath\*" -Destination $parentPath -Force -Recurse -ErrorAction Stop
        
        if ($startService)
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
if ($credential)
{
    $invokeCommandSplat.Credential = $credential
    $invokeCommandSplat.ComputerName = $_machines
}
if ($sessionOption)
{
    $invokeCommandSplat.sessionOption = $sessionOption
}
if ($useSSL)
{
    $invokeCommandSplat.UseSSL = $true
}

Invoke-Command @invokeCommandSplat -ArgumentList $ServiceName, $serviceDisplayName, $serviceDescription, $serviceStartupType, $TimeOut, $StopProcess, $CleanInstall, $ArtifactPath, $installationPath, $startCommand, $runAsCredential, $installTopShelfService, $instanceName, $installArguments, $startService, $RecreateService
Trace-VstsLeavingInvocation $MyInvocation