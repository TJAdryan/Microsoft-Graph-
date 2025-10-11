
# PowerShell Workflow: Certificate-Based Authentication for Azure AD App Registration

This guide walks through creating a certificate, exporting `.pfx` and `.cer` files, uploading the certificate to Azure AD, and granting access to the private key for a specific user.

---

## 🔧 Step 1: Create a Self-Signed Certificate (Exportable)
```powershell
$cert = New-SelfSignedCertificate -CertStoreLocation "cert:\LocalMachine\My" `
    -Subject "CN=SharePoint-Review" `
    -KeyExportPolicy Exportable `
    -KeySpec Signature `
    -KeyLength 2048 `
    -NotAfter (Get-Date).AddYears(2)
```

---

## 📦 Step 2: Export the `.pfx` (includes private key)
```powershell
Export-PfxCertificate -Cert $cert -FilePath "C:\SharePointReviewCert.pfx" `
    -Password (ConvertTo-SecureString -String "YourPassword" -Force -AsPlainText)
```

---

## 📄 Step 3: Export the `.cer` (public key only)
```powershell
$cerBytes = $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert)
[System.IO.File]::WriteAllBytes("C:\SharePointReviewCert.cer", $cerBytes)
```

---

## ☁️ Step 4: Upload `.cer` to Azure AD App Registration
1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory > App registrations**
3. Select your app (Client ID: `350a2-string-and-a51b-etc9ida`)
4. Go to **Certificates & Secrets**
5. Click **Upload Certificate**
6. Upload `C:\SharePointReviewCert.cer`
7. Confirm the thumbprint matches your certificate

---

## 🔐 Step 5: Grant Access to Private Key for Your User
```powershell
$thumbprint = "LETTERSAND###12341243"  # Replace with actual thumbprint
$cert = Get-ChildItem -Path Cert:\LocalMachine\My\$thumbprint
$keyPath = "$env:ProgramData\Microsoft\Crypto\RSA\MachineKeys\" + $cert.PrivateKey.CspKeyContainerInfo.UniqueKeyContainerName

# Grant access to AzureAD\DominickRyan
icacls $keyPath /grant "AzureAD\youknowwhoyouare:F"
```

---

## ✅ Optional: Re-import `.pfx` with PersistKeySet (if needed)
```powershell
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
$cert.Import("C:\SharePointReviewCert.pfx", "YourPassword", "Exportable,PersistKeySet")
$store = New-Object System.Security.Cryptography.X509Certificates.X509Store("My", "LocalMachine")
$store.Open("ReadWrite")
$store.Add($cert)
$store.Close()
```

---

## ✅ Final Notes
- Ensure the certificate thumbprint matches in both PowerShell and Azure.
- Always use `PersistKeySet` to make the private key accessible.
- Grant access to the user running the script if authentication fails with "Keyset does not exist".
