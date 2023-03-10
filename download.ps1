$Menu = {
   Write-Host " *******************************************"
   Write-Host " *                  Menu                   *" 
   Write-Host " *******************************************" 
   Write-Host 
   Write-Host " 1. Download Office 64-bit" 
   Write-Host " 2. Download Office 32-bit"
   Write-Host " 3. Quit"
   Write-Host
   Write-Host
   Write-Host " Information:"
   Write-Host 
   Write-Host "  - Office 32-bit can run on both Windows 10 32 and 64-bit."
   Write-Host "  - Office 64-bit can run on Windows 10 64-bit only."
   Write-Host "  - Nowadays, almost computers run Windows 64-bit."
   Write-Host
   Write-Host " Select an option and press Enter: "  -nonewline
   }
   cls
   
   $Menu1 = {
      Write-Host " *******************************************"
      Write-Host " *                  Menu                   *" 
      Write-Host " *******************************************" 
      Write-Host 
      Write-Host " 1. Microsoft Office 2019" 
      Write-Host " 2. Microsoft Office 2021" 
      Write-Host " 3. Microsoft Office 365" 
      Write-Host " 4. Quit"
      Write-Host 
      Write-Host " Select an option and press Enter: "  -nonewline
   }
   
   $Menu2 = {
      Write-Host " *******************************************"
      Write-Host " *                  Menu                   *" 
      Write-Host " *******************************************" 
      Write-Host 
      Write-Host " 1.  Office $version Professional Plus" 
      Write-Host " 2.  Office $version Standard"
      Write-Host " 3.  Project Professional $version" 
      Write-Host " 4.  Project Standard $version" 
      Write-Host " 5.  Visio Professional $version" 
      Write-Host " 6.  Visio Standard $version" 
      Write-Host " 7.  Word $version" 
      Write-Host " 8.  Excel $version" 
      Write-Host " 9.  PowerPoint $version" 
      Write-Host " 10. Outlook $version" 
      Write-Host " 11. Publisher $version" 
      Write-Host " 12. Access $version" 
      Write-Host " 13. Go Back"
      Write-Host 
      Write-Host " Select an option and press Enter: "  -nonewline
   }

   $Menu365 = {
      Write-Host " ********************************************"
      Write-Host " *                  Menu                    *" 
      Write-Host " ********************************************" 
      Write-Host 
      Write-Host " 1. Office $version Home and Personal"   
      Write-Host " 2. Office $version App for Business"
      Write-Host " 3. Microsoft $version Apps for Enterprise"
      Write-Host " 4. Quit"
      Write-Host
      Write-Host
      Write-Host " Information:"
      Write-Host 
      Write-Host "  - Office 365 Home and Personal using the same package."
      Write-Host "  - Microsoft 365 Apps for Business known as Office 365 Business ."
      Write-Host "  - Microsoft 365 Apps for Enterprise knows as Office 365 ProPlus."
      Write-Host
      Write-Host " Select an option and press Enter: "  -nonewline
   }
   
   
   $download = {
      $null = New-Item -Path "$env:USERPROFILE\Desktop\$productName" -ItemType Directory -Force
      Set-Location "$env:USERPROFILE\Desktop\$productName"
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
      (New-Object Net.WebClient).DownloadFile($uri, "$env:USERPROFILE\Desktop\$productName\setup.exe")

      Write-Host
      Write-Host *****************************************************************************
      Write-Host "- Downloading $productName $($arch) bit...."
      Write-Host "- Once done, it will go back to the menu automatically. Please be patient!"
      Write-Host *****************************************************************************
      Write-Host

      .\setup.exe /download "configuration-x$arch.xml" 
   }

   $download365 = {
      $null = New-Item -Path "$env:USERPROFILE\Desktop\$productName" -ItemType Directory -Force
      Set-Location "$env:USERPROFILE\Desktop\$productName"
      $fileName = "configuration-x$arch.xml"
      $null = New-Item $fileName -ItemType File -Force
      Add-Content $fileName -Value '<Configuration>'
      Add-content $fileName -Value "<Add OfficeClientEdition=`"$arch`" Channel=`"Current`">"
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
      (New-Object Net.WebClient).DownloadFile($uri, "$env:USERPROFILE\Desktop\$productName\setup.exe")
   
      Write-Host
      Write-Host *****************************************************************************
      Write-Host "- Downloading $productName $($arch) bit...."                               
      Write-Host "- Once done, it will go back to the menu automatically. Please be patient!"
      Write-Host *****************************************************************************
      Write-Host
   
      .\setup.exe /download "configuration-x$arch.xml" 
   }

   
   Do { 
   cls
   Invoke-Command $Menu
   $select = Read-Host
   
   if ($select -eq 1) {$arch = '64'}
   if ($select -eq 2) {$arch = '32'}
   
   #Menu level 1 
   Switch ($select)
       {
         # Download Office 64-bit.   
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
      
                        if ($select2 -eq 1) {$productId = "ProPlus$($version)Volume";$productName = "Office $version Professional Plus"}
                        if ($select2 -eq 2) {$productId = "Standard$($version)Volume";$productName = "Office $version Standard"}
                        if ($select2 -eq 3) {$productId = "ProjectPro$($version)Volume";$productName = "Project $version Pro"}
                        if ($select2 -eq 4) {$productId = "ProjectStd$($version)Volume";$productName = "Project $version Standard"}
                        if ($select2 -eq 5) {$productId = "VisioPro$($version)Volume";$productName = "Visio $version Pro"}
                        if ($select2 -eq 6) {$productId = "VisioStd$($version)Volume";$productName = "Visio $version Standard"}
                        if ($select2 -eq 7) {$productId = "Word$($version)Volume";$productName = "Word $version"}
                        if ($select2 -eq 8) {$productId = "Excel$($version)Volume";$productName = "Excel $version"}
                        if ($select2 -eq 9) {$productId = "PowerPoint$($version)Volume";$productName = "PowerPoint $version"}
                        if ($select2 -eq 10) {$productId = "Outlook$($version)Volume";$productName = "Outlook $version"}
                        if ($select2 -eq 11) {$productId = "Publisher$($version)Volume";$productName = "Publisher $version"}
                        if ($select2 -eq 12) {$productId = "Access$($version)Volume";$productName = "Access $version"}
      
                        Switch ($select2)
                           {
                           1 {Invoke-Command $download}
                           2 {Invoke-Command $download}
                           3 {Invoke-Command $download}
                           4 {Invoke-Command $download}
                           5 {Invoke-Command $download}
                           6 {Invoke-Command $download}
                           7 {Invoke-Command $download}
                           8 {Invoke-Command $download}
                           9 {Invoke-Command $download}
                           10 {Invoke-Command $download}
                           11 {Invoke-Command $download}
                           12 {Invoke-Command $download}
      
                           }
                        }
      
                        While ($select2 -ne 13)
                        cls
                     }
      
                     #Download Microsoft Office 2021
                     2 {
                        Do { 
                        cls
                        Invoke-Command $Menu2
                        $select2 = Read-Host
      
                        if ($select2 -eq 1) {$productId = "ProPlus$($version)Volume";$productName = "Office $version Professional Plus LTSC"}
                        if ($select2 -eq 2) {$productId = "Standard$($version)Volume";$productName = "Office $version Standard LTSC"}
                        if ($select2 -eq 3) {$productId = "ProjectPro$($version)Volume";$productName = "Project $version Pro LTSC"}
                        if ($select2 -eq 4) {$productId = "ProjectStd$($version)Volume";$productName = "Project $version Standard LTSC"}
                        if ($select2 -eq 5) {$productId = "VisioPro$($version)Volume";$productName = "Visio $version Pro LTSC"}
                        if ($select2 -eq 6) {$productId = "VisioStd$($version)Volume";$productName = "Visio $version Standard LTSC"}
                        if ($select2 -eq 7) {$productId = "Word$($version)Volume";$productName = "Word $version LTSC"}
                        if ($select2 -eq 8) {$productId = "Excel$($version)Volume";$productName = "Excel $version LTSC"}
                        if ($select2 -eq 9) {$productId = "PowerPoint$($version)Volume";$productName = "PowerPoint $version LTSC"}
                        if ($select2 -eq 10) {$productId = "Outlook$($version)Volume";$productName = "Outlook $version LTSC"}
                        if ($select2 -eq 11) {$productId = "Publisher$($version)Volume";$productName = "Publisher $version LTSC"}
                        if ($select2 -eq 12) {$productId = "Access$($version)Volume";$productName = "Access $version LTSC"}
      
                        Switch ($select2)
                           {
                              1 {Invoke-Command $download}
                              2 {Invoke-Command $download}
                              3 {Invoke-Command $download}
                              4 {Invoke-Command $download}
                              5 {Invoke-Command $download}
                              6 {Invoke-Command $download}
                              7 {Invoke-Command $download}
                              8 {Invoke-Command $download}
                              9 {Invoke-Command $download}
                              10 {Invoke-Command $download}
                              11 {Invoke-Command $download}
                              12 {Invoke-Command $download}
      
                           }
                        }
      
                        While ($select2 -ne 13)
                        cls
                     }
      
                     #Download Microsoft Office 365
                     3 {
                        Do { 
                        cls
                        Invoke-Command $Menu365
                        $select2 = Read-Host
      
                        if ($select2 -eq 1) {$productId = "O365HomePremRetail";$productName = "Microsoft $version Home & Personal"}
                        if ($select2 -eq 2) {$productId = "O365BusinessRetail";$productName = "Microsoft $version Apps for Business"}
                        if ($select2 -eq 3) {$productId = "O365ProPlusRetail";$productName = "Microsoft $version Apps for Enterprise"}
      
                        Switch ($select2)
                           {
                           1 {Invoke-Command $download365}
                           2 {Invoke-Command $download365}
                           3 {Invoke-Command $download365}
      
                           }
                        }
      
                        While ($select2 -ne 4)
                     }


      
                  }
               }
      
               While ($select1 -ne 4)
               cls
      
               }


         # Download Office 32-bit.   
         2 {
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

                     if ($select2 -eq 1) {$productId = "ProPlus$($version)Volume";$productName = "Office $version Professional Plus"}
                     if ($select2 -eq 2) {$productId = "Standard$($version)Volume";$productName = "Office $version Standard"}
                     if ($select2 -eq 3) {$productId = "ProjectPro$($version)Volume";$productName = "Project $version Pro"}
                     if ($select2 -eq 4) {$productId = "ProjectStd$($version)Volume";$productName = "Project $version Standard"}
                     if ($select2 -eq 5) {$productId = "VisioPro$($version)Volume";$productName = "Visio $version Pro"}
                     if ($select2 -eq 6) {$productId = "VisioStd$($version)Volume";$productName = "Visio $version Standard"}
                     if ($select2 -eq 7) {$productId = "Word$($version)Volume";$productName = "Word $version"}
                     if ($select2 -eq 8) {$productId = "Excel$($version)Volume";$productName = "Excel $version"}
                     if ($select2 -eq 9) {$productId = "PowerPoint$($version)Volume";$productName = "PowerPoint $version"}
                     if ($select2 -eq 10) {$productId = "Outlook$($version)Volume";$productName = "Outlook $version"}
                     if ($select2 -eq 11) {$productId = "Publisher$($version)Volume";$productName = "Publisher $version"}
                     if ($select2 -eq 12) {$productId = "Access$($version)Volume";$productName = "Access $version"}

                     Switch ($select2)
                        {
                        1 {Invoke-Command $download}
                        2 {Invoke-Command $download}
                        3 {Invoke-Command $download}
                        4 {Invoke-Command $download}
                        5 {Invoke-Command $download}
                        6 {Invoke-Command $download}
                        7 {Invoke-Command $download}
                        8 {Invoke-Command $download}
                        9 {Invoke-Command $download}
                        10 {Invoke-Command $download}
                        11 {Invoke-Command $download}
                        12 {Invoke-Command $download}
   
                        }
                     }

                     While ($select2 -ne 13)
                     cls
                  }

                  #Download Microsoft Office 2021
                  2 {
                     Do { 
                     cls
                     Invoke-Command $Menu2
                     $select2 = Read-Host

                     if ($select2 -eq 1) {$productId = "ProPlus$($version)Volume";$productName = "Office $version Professional Plus LTSC"}
                     if ($select2 -eq 2) {$productId = "Standard$($version)Volume";$productName = "Office $version Standard LTSC"}
                     if ($select2 -eq 3) {$productId = "ProjectPro$($version)Volume";$productName = "Project $version Pro LTSC"}
                     if ($select2 -eq 4) {$productId = "ProjectStd$($version)Volume";$productName = "Project $version Standard LTSC"}
                     if ($select2 -eq 5) {$productId = "VisioPro$($version)Volume";$productName = "Visio $version Pro LTSC"}
                     if ($select2 -eq 6) {$productId = "VisioStd$($version)Volume";$productName = "Visio $version Standard LTSC"}
                     if ($select2 -eq 7) {$productId = "Word$($version)Volume";$productName = "Word $version LTSC"}
                     if ($select2 -eq 8) {$productId = "Excel$($version)Volume";$productName = "Excel $version LTSC"}
                     if ($select2 -eq 9) {$productId = "PowerPoint$($version)Volume";$productName = "PowerPoint $version LTSC"}
                     if ($select2 -eq 10) {$productId = "Outlook$($version)Volume";$productName = "Outlook $version LTSC"}
                     if ($select2 -eq 11) {$productId = "Publisher$($version)Volume";$productName = "Publisher $version LTSC"}
                     if ($select2 -eq 12) {$productId = "Access$($version)Volume";$productName = "Access $version LTSC"}

                     Switch ($select2)
                        {
                           1 {Invoke-Command $download}
                           2 {Invoke-Command $download}
                           3 {Invoke-Command $download}
                           4 {Invoke-Command $download}
                           5 {Invoke-Command $download}
                           6 {Invoke-Command $download}
                           7 {Invoke-Command $download}
                           8 {Invoke-Command $download}
                           9 {Invoke-Command $download}
                           10 {Invoke-Command $download}
                           11 {Invoke-Command $download}
                           12 {Invoke-Command $download}
   
                        }
                     }

                     While ($select2 -ne 13)
                     cls
                  }

                  #Download Microsoft Office 365
                  3 {
                     Do { 
                     cls
                     Invoke-Command $Menu365
                     $select2 = Read-Host

                     if ($select2 -eq 1) {$productId = "O365HomePremRetail";$productName = "Microsoft $version Home & Personal"}
                     if ($select2 -eq 2) {$productId = "O365BusinessRetail";$productName = "Microsoft $version Apps for Business"}
                     if ($select2 -eq 3) {$productId = "O365ProPlusRetail";$productName = "Microsoft $version Apps for Enterprise"}

                     Switch ($select2)
                        {
                        1 {Invoke-Command $download365}
                        2 {Invoke-Command $download365}
                        3 {Invoke-Command $download365}
   
                        }
                     }

                     While ($select2 -ne 4)
                     cls
                  }



               }
            }

            While ($select1 -ne 4)
            cls

            }


       }
   }
   
   While ($select -ne 3)
   cls
