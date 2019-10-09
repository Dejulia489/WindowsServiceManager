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
