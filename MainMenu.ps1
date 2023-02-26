$MainMenu = {
Write-Host " *******************************************"
Write-Host " *                  Menu                   *" 
Write-Host " *******************************************" 
Write-Host 
Write-Host " 1. Oprion 1" 
Write-Host " 2. Oprion 2" 
Write-Host " 3. Quit"
Write-Host 
Write-Host " Select an option and press Enter: "  -nonewline
}
cls

$MainMenu1 = {
Write-Host " *******************************************"
Write-Host " *                  Menu                   *" 
Write-Host " *******************************************" 
Write-Host 
Write-Host " 1. Oprion 11" 
Write-Host " 2. Oprion 12" 
Write-Host " 3. Quit"
Write-Host 
Write-Host " Select an option and press Enter: "  -nonewline
}

$MainMenu2 = {
Write-Host " *******************************************"
Write-Host " *                  Menu                   *" 
Write-Host " *******************************************" 
Write-Host 
Write-Host " 1. Oprion 112" 
Write-Host " 2. Oprion 122" 
Write-Host " 3. Quit"
Write-Host 
Write-Host " Select an option and press Enter: "  -nonewline
}

$ofin = {
$null = New-Item -Path $env:temp\c2r -ItemType Directory -Force
Set-Location $env:temp\c2r
$fileName = 'configuration.xml'
New-Item $fileName -ItemType File -Force | Out-Null
Add-Content $fileName -Value '<Configuration>'
Add-content $fileName -Value '<Add OfficeClientEdition="64">'
Add-content $fileName -Value "<Product ID=`"$productId`">"
Add-content $fileName -Value '<Language ID="en-us" />'
Add-Content $fileName -Value '</Product>'
Add-Content $fileName -Value '</Add>'
Add-Content $fileName -Value '</Configuration>'
Write-Host
Write-Host ============================================================
Write-Host "Installing $productId...."
Write-Host ============================================================
Write-Host

$uri = 'https://github.com/bonben365/office365-installer/raw/main/setup.exe'
Invoke-WebRequest -Uri $uri -OutFile 'setup.exe' -ErrorAction:SilentlyContinue | Out-Null
.\setup.exe /configure .\configuration.xml


}

Do { 
cls
Invoke-Command $MainMenu
$Select = Read-Host
Switch ($Select)
    {
    1 {
         cls
         Do { 
         cls
         Invoke-Command $MainMenu1
         $Select1 = Read-Host
         Switch ($Select1)
            {
               1 {
                  Do { 
                  cls
                  Invoke-Command $MainMenu2
                  $Select2 = Read-Host

                  Switch ($Select2)
                     {
                     1 {Invoke-Command $ofin}
                     2 {Invoke-Command $ofin}
                     3 {Invoke-Command $ofin}
                     4 {Invoke-Command $ofin}
 
                     }
                  }
                  While ($Select -ne 3)
               }


            }
         }

         While ($Select1 -ne 3)



       }
    2 {
       #Insert your code here

       }
    }
}

While ($Select -ne 3)