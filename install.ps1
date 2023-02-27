$Menu = {
   Write-Host " *******************************************"
   Write-Host " *                  Menu                   *" 
   Write-Host " *******************************************" 
   Write-Host 
   Write-Host " 1. Download Office 64-bit" 
   Write-Host " 2. Download Office 32-bit" 
   Write-Host " 3. Quit"
   Write-Host 
   Write-Host " select an option and press Enter: "  -nonewline
   }
   cls
   
   $Menu1 = {
   Write-Host " *******************************************"
   Write-Host " *                  Menu                   *" 
   Write-Host " *******************************************" 
   Write-Host 
   Write-Host " 1. Office 2019" 
   Write-Host " 2. Office 2021" 
   Write-Host " 2. Office 365" 
   Write-Host " 3. Quit"
   Write-Host 
   Write-Host " select an option and press Enter: "  -nonewline
   }
   
   $Menu2 = {
   Write-Host " *******************************************"
   Write-Host " *                  Menu                   *" 
   Write-Host " *******************************************" 
   Write-Host 
   Write-Host " 1. Office 2019 Professional Plus" 
   Write-Host " 2. Office 2019 Standard" 
   Write-Host " 3. Quit"
   Write-Host 
   Write-Host " select an option and press Enter: "  -nonewline
   }
   
   
   $install = {
   $null = New-Item -Path "~\Desktop\Office$version" -ItemType Directory -Force
   Set-Location "~\Desktop\Office$version"
   $fileName = "configuration-x$arch.xml"
   $null = New-Item $fileName -ItemType File -Force
   Add-Content $fileName -Value '<Configuration>'
   Add-content $fileName -Value "<Add OfficeClientEdition=`"$arch`">"
   Add-content $fileName -Value "<Product ID=`"$productId`">"
   Add-content $fileName -Value '<Language ID="en-us" />'
   Add-Content $fileName -Value '</Product>'
   Add-Content $fileName -Value '</Add>'
   Add-Content $fileName -Value '</Configuration>'
   
   #Create the installation batch files
   $batchName = "Install-x$arch.bat"
   $null = New-Item $batchName -ItemType File -Force
   Add-Content $batchName -Value "setup.exe /configure configuration-x$arch.xml"
   
   # Download the Office Deployment Tool
   $uri = 'https://github.com/bonben365/office365-installer/raw/main/setup.exe'
   $null = Invoke-WebRequest -Uri $uri -OutFile 'setup.exe' -ErrorAction:SilentlyContinue
   .\setup.exe /download "configuration-x$arch.xml"
   
   Write-Host
   Write-Host ***************************************************************
   Write-Host "Downloading $productName $($arch) bit...."                    *
   Write-Host ***************************************************************
   Write-Host
   
   }
   
   Do { 
   cls
   Invoke-Command $Menu
   $select = Read-Host
   
   if ($select -eq 1) {$arch = '64'}
   if ($select -eq 2) {$arch = '86'}
   
   #Menu level 1 
   Switch ($select)
       { 
       1 {
            cls
   
            Do { 
            cls
            Invoke-Command $Menu1
            $select1 = Read-Host
   
            if ($select1 -eq 1) {$version = '2019'}
            if ($select1 -eq 2) {$version = '2021'}
            if ($select1 -eq 3) {$version = '365'}
   
            #Menu level 2
            Switch ($select1)
               {
                  #Download Microsoft Office 2019
                  1 {
                     Do { 
                     cls
                     Invoke-Command $Menu2
                     $select2 = Read-Host
   
                     if ($select2 -eq 1) {$productId = 'ProPlus2019Volume';$productName = 'Office 2019 Professional Plus'}
                     if ($select2 -eq 2) {$productId = 'Standard2019Volume'}
                     if ($select2 -eq 3) {$productId = 'ProPlus2019Volume'}
                     if ($select2 -eq 4) {$productId = 'ProPlus2019Volume'}
   
                     Switch ($select2)
                        {
                        1 {Invoke-Command $install}
                        2 {Invoke-Command $install}
                        3 {Invoke-Command $install}
                        4 {Invoke-Command $install}
    
                        }
                     }
   
                     While ($select -ne 5)
                  }
   
                  #Download Microsoft Office 2021
                  2 {
                     Do { 
                     cls
                     Invoke-Command $Menu2
                     $select2 = Read-Host
   
                     if ($select2 -eq 1) {$productId = 'ProPlus2021Volume'}
                     if ($select2 -eq 2) {$productId = 'Standard2021Volume'}
   
                     #Download Microsoft Office 2019
                     Switch ($select2)
                        {
                        1 {Invoke-Command $install}
                        2 {Invoke-Command $install}
                        3 {Invoke-Command $install}
                        4 {Invoke-Command $install}
    
                        }
                     }
                     
                     While ($select -ne 3)
                  }
   
   
               }
            }
   
            While ($select1 -ne 4)
   
   
   
            }
       2 {}
       3 {}
       }
   }
   
   While ($select -ne 3)