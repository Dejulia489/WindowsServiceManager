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
            Write-Error "Service responded with [$($respone.ReturnValue)], expected 0" -ErrorAction Stop
        }
    }
    $parentPath = $serviceObject.PathName | Split-Path -Parent
    Write-Output "Identified deployment location [$parentPath]"
    If ($parentPath)
    {
        If ($CleanInstall)
        {
            Write-Output "[$($MyInvocation.MyCommand.Name)]: Clean install requested, removing [$parentPath]"
            Remove-Item -Path $parentPath -Force -Recurse
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
        Write-Error "Service responded with [$($respone.ReturnValue)], expected 0" -ErrorAction Stop
    }
}
else
{
    Write-Error "Unable to locate [$ServiceName] on [$Env:ComputerName], confirm the service is installed correctly." -ErrorAction Stop   
}
Trace-VstsLeavingInvocation $MyInvocation