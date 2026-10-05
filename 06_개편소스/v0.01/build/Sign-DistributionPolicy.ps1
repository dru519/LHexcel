param(
    [Parameter(Mandatory=$true)][ValidatePattern('^NXL1\|v0\.01_r[0-9]+\|(internal-xlam|enhanced-dll)\|[0-9]{4}-[0-9]{2}-[0-9]{2}$')][string]$Payload,
    [string]$KeyContainer = 'LHExcel.DistributionPolicy.v1'
)
$ErrorActionPreference = 'Stop'
# The private key stays in this Windows user's CSP key store. Never export it.
$parameters = New-Object System.Security.Cryptography.CspParameters(24)
$parameters.KeyContainerName = $KeyContainer
$parameters.KeyNumber = 2
$parameters.Flags = [System.Security.Cryptography.CspProviderFlags]::UseNonExportableKey
$rsa = New-Object System.Security.Cryptography.RSACryptoServiceProvider(3072, $parameters)
try {
    $rsa.PersistKeyInCsp = $true
    $bytes = [Text.Encoding]::ASCII.GetBytes($Payload)
    $signature = $rsa.SignData($bytes, 'SHA256')
    if (-not $rsa.VerifyData($bytes, 'SHA256', $signature)) { throw 'Policy signature self-check failed' }
    # CryptoAPI consumes little-endian RSA signatures; .NET returns big-endian.
    [Array]::Reverse($signature)
    @{
        payload=$Payload
        public_blob_hex=([BitConverter]::ToString($rsa.ExportCspBlob($false))).Replace('-','').ToLowerInvariant()
        signature_hex=([BitConverter]::ToString($signature)).Replace('-','').ToLowerInvariant()
        algorithm='RSA-3072-SHA256-PKCS1'
        publisher_certificate=$false
    } | ConvertTo-Json -Compress
} finally { $rsa.Dispose() }
