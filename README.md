# Outlook signature generator

Simple command line utility that generates Outlook signature based on Word template.


## Usage 

```
Mail.OutlookSignature.exe "signature-template.docx"
```

### App configuration parameters

SignatureName
LockSignature
LockSignatureOverrideGroupName

## How it works
When app is started it accepts one parameter that should contains path to a Word document with template variables that would be replaced with actual values from Active Directory of currently logged on user.

It replaces parameters, sets this document as default signature and locks registry to block changes to the signature by end user (to force using corporate template by everyone).

If Outlook was running while running the tool, Outlook application needs to be restarted to see the signature.

Generated signature is stored in `%appdata%\Microsoft\Signatures` folder.

## Supported variables in template

| Variable                        | Source LDAP field            | Description         | ADUC Tab     |
| ------------------------------- | ---------------------------- | ------------------- | ------------ |
|  `%givenName%`                  | `givenName`                  | First name          | General      |
|  `%sn%`                         | `sn`                         | Last name           | General      |
|  `%displayName%`                | `displayName`                | Display name        | General      |
|  `%department%`                 | `department`                 | Department          | Organization |
|  `%company%`                    | `company`                    | Company             | Organization |
|  `%telephoneNumber%`            | `telephoneNumber`            | Telephone number    | General      |
|  `%mobile%`                     | `mobile`                     | Mobile              | Telephones   |
|  `%mail%`                       | `mail`                       | E-Mail              | General      |
|  `%physicalDeliveryOfficeName%` | `physicalDeliveryOfficeName` | Office              | General      |
|  `%postalCode%`                 | `postalCode`                 | Zip/Postal Code     | Address      |
|  `%streetAddress%`              | `streetAddress`              | Street              | Address      |
|  `%title%`                      | `title`                      | Job Title           | Organization |
|  `%l%`                          | `l`                          | City                | Address      |
|  `%st%`                         | `st`                         | State/province      | Address      |
|  `%sc%`                         | `c`                          | Country             | Address      |
| `%country%`                     |                              | Expanded country name using internal dictionary | -- |
| `%QR%`                          |                              | QR Code with VCARD content | -- |

### Example

### Word template document
![Word template](docs/word-template.png | width=200)

### Group policy Logon script
Set-Signature.ps1
```powershell
Start-Process -FilePath "$($PSScriptRoot)\App\Mail.OutlookSignature.exe" `
    -WorkingDirectory $PSScriptRoot -ArgumentList (Join-Path $PSScriptRoot "template.docx") `
    -NoNewWindow `
    -Wait
``` 

![Logon script](docs/logon-script.png | width=200)

### Generated signature in Outlook
![Generated signature](docs/generated-signature.png | width=200)

## Dependencies

- https://github.com/OfficeDev/Open-Xml-PowerTools for Word document parsing
- https://www.nuget.org/packages/MessagingToolkit.QRCode for QR code image generation
