
# Download Offline Installer for all Microsoft Office Apps

## Information
- All downloaded files are saved in your Desktop.
- Do not turn off or disconnected from the internet.
- You can download 32 or 64 bit version of Office as you need.
- The script support download Office 2019, Office 2021 and Microsoft 365 Apps (Home, Business and Enterprise)
- Download single app such as Word, Excel, PowerPoint... is supported also.
- Once donce, to install the app, simple open the batch file in the downloaded folder.

  
## Installation

Right click on the Windows Start icon 🪟 then select Windows PowerShell (Admin). 
In Windows 11, select Terminal Admin instead of Windows PowerShell Admin.

![App Screenshot](https://s3.amazonaws.com/s3.bonben365.com/files/cons/powershell10.jpg)

Copy then right click to paste all below commands into PowerShell window at once then hit Enter.

```ps
Set-ExecutionPolicy Bypass -Scope Process -Force
$url='https://github.com/bonben365/office-downloader/raw/main/download.ps1'
iex ((New-Object System.Net.WebClient).DownloadString($url))
```

![App Screenshot](https://s3.amazonaws.com/s3.bonben365.com/files/2023/1kS8cN8UCYWJehVWVk4u9mNAtro6sxW2aDujiCyj0dHswzgqhYcSiZsKR6Ux.jpg)

![App Screenshot](https://s3.amazonaws.com/s3.bonben365.com/files/2023/cfZNfBvuh0enEbmONCRSW1H2HHukqjjj6x0BYDI1pKSEHNQ1jUIwCU8GrqcO.jpg)

![App Screenshot](https://s3.amazonaws.com/s3.bonben365.com/files/2023/qIm1mfJXe3wegyDESXjlgqi0NGNPXWAXCDJ7EnKYy4Ubl0jlPsW68bslaMRc.jpg)

![App Screenshot](https://s3.amazonaws.com/s3.bonben365.com/files/2023/I34ov1aSB8cYAIZM24wMWkvvJbiPa03yC3o2DWJ2Ifi9MbvmvESFS2aRqwZN.jpg)

Once done, open the Install.bat file to install the app

![App Screenshot](https://s3.amazonaws.com/s3.bonben365.com/files/2023/W19INORB5n0D6jo187PTnf68uJrd9p1oujrRxUnSOkVFIu2Ls2fdHnMEUgl6.jpg)


➡️Please inspect [https://github.com/bonben365/office-downloader/raw/main/download.ps1](https://github.com/bonben365/office-downloader/raw/main/download.ps1) prior to running any of these scripts to ensure safety. We already know it's safe, but you should verify the security and contents of any script from the internet you are not familiar with.

## Documentation

[🌍https://bonben365.com/](https://bonben365.com/)

