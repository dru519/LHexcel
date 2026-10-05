param([Parameter(Mandatory=$true)][string]$X86Dll,[Parameter(Mandatory=$true)][string]$X64Dll)
& (Join-Path $PSScriptRoot 'Register-NxHost.ps1') -Action Unregister -X86Dll $X86Dll -X64Dll $X64Dll
