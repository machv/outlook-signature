# Outlook signature generator

Simple command line utility that generates Outlook signature based on Word template.


## Usage 

```
Mail.OutlookSignature.exe "signature-template.docx"

```

## How it works
When app is started it accepts one parameter that should contains path to a Word document with template variables that would be replaced with actual values from Active Directory of currently logged on user.

It replaces parameters, sets this document as default signature and locks registry to block changes to the signature by end user (to force using corporate template by everyone).

If Outlook was running while running the tool, Outlook application needs to be restarted to see the signature.
