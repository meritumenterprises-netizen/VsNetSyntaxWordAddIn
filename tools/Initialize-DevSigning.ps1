# Run with PowerShell 7. The development certificate stays local to this project.
$ErrorActionPreference = 'Stop'
$certificatePath = Join-Path (Split-Path $PSScriptRoot -Parent) 'DevelopmentSigning.pfx'
if (Test-Path -LiteralPath $certificatePath) {
    Write-Host 'Development signing certificate already exists.'
} else {
$rsa = [System.Security.Cryptography.RSA]::Create(2048)
try {
    $request = [System.Security.Cryptography.X509Certificates.CertificateRequest]::new(
        'CN=VS.NET Syntax Local Development',
        $rsa,
        [System.Security.Cryptography.HashAlgorithmName]::SHA256,
        [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
    $oids = [System.Security.Cryptography.OidCollection]::new()
    [void]$oids.Add([System.Security.Cryptography.Oid]::new('1.3.6.1.5.5.7.3.3'))
    $request.CertificateExtensions.Add(
        [System.Security.Cryptography.X509Certificates.X509EnhancedKeyUsageExtension]::new($oids, $false))
    $certificate = $request.CreateSelfSigned([DateTimeOffset]::Now.AddDays(-1), [DateTimeOffset]::Now.AddYears(1))
    try {
        [System.IO.File]::WriteAllBytes($certificatePath,
            $certificate.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Pfx, ''))
    } finally { $certificate.Dispose() }
} finally { $rsa.Dispose() }
Write-Host "Created local development certificate: $certificatePath"
}
$signingCertificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new(
    $certificatePath, '', [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::EphemeralKeySet)
try {
    $settings = [System.Xml.XmlWriterSettings]::new()
    $settings.Indent = $true
    $writer = [System.Xml.XmlWriter]::Create((Join-Path (Split-Path $PSScriptRoot -Parent) 'DevelopmentSigning.props'), $settings)
    try {
        $writer.WriteStartElement('Project', 'http://schemas.microsoft.com/developer/msbuild/2003')
        $writer.WriteStartElement('PropertyGroup')
        $writer.WriteElementString('ManifestCertificateThumbprint', $signingCertificate.Thumbprint)
        $writer.WriteEndElement()
        $writer.WriteEndElement()
    } finally { $writer.Dispose() }
} finally { $signingCertificate.Dispose() }
