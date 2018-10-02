param
(
    [Parameter(Mandatory)]
    [string]
    $ServiceName = (Get-VstsInput -Name 'ServiceName' -Require), 

    [Parameter(Mandatory)]
    [string]
    $PackagePath = (Get-VstsInput -Name 'PackagePath' -Require),

    [Parameter()]
    [string]
    $CleanInstall = (Get-VstsInput -Name 'CleanInstall' -AsBool)
)

Write-Output "Getting [$ServiceName]"
$serviceObject = Get-WmiObject -Class Win32_Service | Where-Object {$PSItem.Name -eq $ServiceName}
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
If ($parentPath)
{
    If ($CleanInstall)
    {
        Write-Output "[$($MyInvocation.MyCommand.Name)]: Clean install requested, removing [$parentPath]"
        Remove-Item -Path $parentPath -Force -Recurse
    }
}
Else
{
    $null = New-Item -ItemType Directory -Path $parentPath -Force
}
Write-Output "Copying [$PackagePath] to [$parentPath]"
Copy-Item -Path $PackagePath\* -Destination $parentPath -Force -Recurse
Write-Output "Starting [$ServiceName]"
$respone = $serviceObject.StartService()
If ($respone.ReturnValue -ne 0)
{
    Write-Error "Service responded with [$($respone.ReturnValue)], expected 0" -ErrorAction Stop
}
