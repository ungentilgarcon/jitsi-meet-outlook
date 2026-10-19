[CmdletBinding()]
param(
    [string] $CertificatePath = (Join-Path $PSScriptRoot 'Visiovilleurbanne.cer')
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $CertificatePath -PathType Leaf)) {
    throw "Certificate file not found: $CertificatePath"
}

$resolvedPath = (Resolve-Path -LiteralPath $CertificatePath).Path
$certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($resolvedPath)

try {
    $expectedThumbprint = '6AA2E9073B220D37396D87942FFA5E54B349294D'
    if ($certificate.Thumbprint -ne $expectedThumbprint -or
        $certificate.Subject -ne 'CN=Visio Villeurbanne local test') {
        throw "The certificate does not match the expected Visio Villeurbanne certificate. Subject: $($certificate.Subject); thumbprint: $($certificate.Thumbprint)"
    }

    if ($certificate.NotBefore -gt (Get-Date) -or $certificate.NotAfter -le (Get-Date)) {
        throw "The certificate is not currently valid. Validity: $($certificate.NotBefore) through $($certificate.NotAfter)"
    }

    $ekuExtension = $certificate.Extensions |
        Where-Object { $_.Oid.Value -eq '2.5.29.37' } |
        Select-Object -First 1
    $hasCodeSigningUsage = $ekuExtension -and
        ($ekuExtension.EnhancedKeyUsages | Where-Object { $_.Value -eq '1.3.6.1.5.5.7.3.3' })
    if (-not $hasCodeSigningUsage) {
        throw 'The certificate is not marked for code signing.'
    }

    $storeLocations = @(
        'Cert:\CurrentUser\TrustedPublisher',
        'Cert:\CurrentUser\Root'
    )
    foreach ($storeLocation in $storeLocations) {
        $alreadyInstalled = Get-ChildItem -Path $storeLocation |
            Where-Object { $_.Thumbprint -eq $expectedThumbprint } |
            Select-Object -First 1

        if ($alreadyInstalled) {
            Write-Output "Already present in $storeLocation"
        }
        else {
            Import-Certificate -FilePath $resolvedPath -CertStoreLocation $storeLocation |
                Out-Null
            Write-Output "Imported into $storeLocation"
        }
    }
}
finally {
    $certificate.Dispose()
}
