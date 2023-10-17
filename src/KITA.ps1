# KITA 2.0

function menu {
 Write-Output "###########################################"
 Write-Output "#                                         #"
 Write-Output "#                KITA 1.0                 #"
 Write-Output "#                                         #"
 Write-Output "#       1.- Install MS Office             #"
 Write-Output "#       2.- Activate Office(KMS)          #"
 Write-Output "#       3.- Remove KMS activation         #"
 #Write-Output "#       4.- Uninstall MS Office           #"
 Write-Output "#       4.- Exit                          #"
 Write-Output "#                                         #"
 Write-Output "###########################################"
$choise= Read-Host "`n[?] Select option "
    while ($choise -ne "1" -and $choise -ne "2" -and $choise -ne "3" -and $choise -ne "4") {
      Write-Host "[*] Please select valid option"
      $choise= Read-Host "[?] Select option "
    }    
    if($choise -eq "1"){officeOpts; Office_Deploy; postInstall}
    if($choise -eq "2"){JAMSOA}
    if($choise -eq "3"){
      Write-Output "[*] For problems with remove KMS Activation check my blog: https://quantumwavves.github.io"; RMSOA }
    #if($choise -eq "4"){removeOffice} soon feature 
    if($choise -eq "4"){Write-Output "[*] Good Bye"}
}

function officeOpts {
 Write-Output "`n###########################################"
 Write-Output "#                                         #"
 Write-Output "#            Office Version               #"
 Write-Output "#                                         #"
 Write-Output "#           1.- 2019 ProPlus              #"
 Write-Output "#           2.- 2021 LTSC                 #"
 Write-Output "#           3.- Office 365                #"
 Write-Output "#                                         #"
 Write-Output "###########################################"
 $versionChoise = Read-Host "`n[?] Select your office version"
 while ($versionChoise -ne 1 -and $versionChoise -ne 2 -and $versionChoise -ne 3) {
  Write-Output "[*] Invalid option, please select valid option"
  $versionChoise = Read-Host "[?] Select option"
 }
 if ($versionChoise -eq "1") {$version = "2019"}
 if ($versionChoise -eq "2") {$version = "2021"}
 if ($versionChoise -eq "3") {$version = "365"}
 Write-Output "`n######################################################################################################################"
 Write-Output "#                                                                                                                    #"
 Write-Output "#                                          Select bloat version                                                      #"
 Write-Output "#                                                                                                                    #"
 Write-Output "#       1.- Meta (Word, Excel, PowerPoint, OneNote)                                                                  #"
 Write-Output "#       2.- Minimal (Word, Excel, PowerPoint, OneNote, Teams)                                                        #"
 Write-Output "#       3.- Normal (Word, Excel, PowerPoint, Teams, Outlook, OneNote, OneDrive)                                      #"
 Write-Output "#       4.- All (Word, Excel, PowerPoint, Teams, Outlook, OneNote, OneDrive, Access, Visio, Publisher, Skype)        #"
 Write-Output "#                                                                                                                    #"
 Write-Output "######################################################################################################################"
 $bloatChoise = Read-Host "`n[?] Select your bloat"
 while ($bloatChoise -ne 1 -and $bloatChoise -ne 2 -and $bloatChoise -ne 3 -and $bloatChoise -ne 4) {
  Write-Output "[*] Invalid option, please select valid option"
  $bloatChoise = Read-Host "[?] Select option"
 }
 if ($bloatChoise -eq "1") {$bloatVersion = "meta"}
 if ($bloatChoise -eq "2") {$bloatVersion = "minimal"}
 if ($bloatChoise -eq "3") {$bloatVersion = "normal"}
 if ($bloatChoise -eq "4") {$bloatVersion = "all"}
 xmlDownloader
}

function xmlDownloader {
  $Path = $env:temp
  $wc = (New-Object System.Net.WebClient)
 
  if ($version -eq "2019"){
    if ($bloatVersion -eq "meta"){
      Write-Output "`n[-] Donwload XML Office 2019 $bloatVersion configuration for deploy"
      $wc.DownloadFile("https://raw.githubusercontent.com/quantumwavves/KITA/master/resources/assets/2019/2019-meta.xml", "$Path\MSOffice.xml")
      }
    if ($bloatVersion -eq "minimal"){
      Write-Output "`n[-] Donwload XML Office 2019 $bloatVersion configuration for deploy"
      $wc.DownloadFile("https://raw.githubusercontent.com/quantumwavves/KITA/master/resources/assets/2019/2019-minimal.xml", "$Path\MSOffice.xml")
      }
    if ($bloatVersion -eq "normal"){
      Write-Output "`n[-] Donwload XML Office 2019 $bloatVersion configuration for deploy"
      $wc.DownloadFile("https://raw.githubusercontent.com/quantumwavves/KITA/master/resources/assets/2019/2019-normal.xml", "$Path\MSOffice.xml")
      }
    if ($bloatVersion -eq "all"){
      Write-Output "`n[-] Donwload XML Office 2019 $bloatVersion configuration for deploy"
      $wc.DownloadFile("https://raw.githubusercontent.com/quantumwavves/KITA/master/resources/assets/2019/2019-full.xml", "$Path\MSOffice.xml")
      }
  }
  if ($version -eq "2021"){
    if ($bloatVersion -eq "meta"){
      Write-Output "`n[-] Donwload XML Office 2021 $bloatVersion configuration for deploy"
      $wc.DownloadFile("https://raw.githubusercontent.com/quantumwavves/KITA/master/resources/assets/2021/2021-meta.xml", "$Path\MSOffice.xml")
      }
    if ($bloatVersion -eq "minimal"){
      Write-Output "`n[-] Donwload XML Office 2021 $bloatVersion configuration for deploy"
      $wc.DownloadFile("https://raw.githubusercontent.com/quantumwavves/KITA/master/resources/assets/2021/2021-minimal.xml", "$Path\MSOffice.xml")
      }
    if ($bloatVersion -eq "normal"){
      Write-Output "`n[-] Donwload XML Office 2021 $bloatVersion configuration for deploy"
      $wc.DownloadFile("https://raw.githubusercontent.com/quantumwavves/KITA/master/resources/assets/2021/2021-normal.xml", "$Path\MSOffice.xml")
      }
    if ($bloatVersion -eq "all"){
      Write-Output "`n[-] Donwload XML Office 2021 $bloatVersion configuration for deploy"
      $wc.DownloadFile("https://raw.githubusercontent.com/quantumwavves/KITA/master/resources/assets/2021/2021-full.xml", "$Path\MSOffice.xml")
      }
  }
  if ($version -eq "365"){
    if ($bloatVersion -eq "meta"){
      Write-Output "`n[-] Donwload XML Office 365 $bloatVersion configuration for deploy"
      $wc.DownloadFile("https://raw.githubusercontent.com/quantumwavves/KITA/master/resources/assets/365/365-meta.xml", "$Path\MSOffice.xml")
      }
    if ($bloatVersion -eq "minimal"){
      Write-Output "`n[-] Donwload XML Office 365 $bloatVersion configuration for deploy"
      $wc.DownloadFile("https://raw.githubusercontent.com/quantumwavves/KITA/master/resources/assets/365/365-minimal.xml", "$Path\MSOffice.xml")
      }
    if ($bloatVersion -eq "normal"){
      Write-Output "`n[-] Donwload XML Office 365 $bloatVersion configuration for deploy"
      $wc.DownloadFile("https://raw.githubusercontent.com/quantumwavves/KITA/master/resources/assets/365/365-normal.xml", "$Path\MSOffice.xml")
      }
    if ($bloatVersion -eq "all"){
      Write-Output "`n[-] Donwload XML Office 365 $bloatVersion configuration for deploy"
      $wc.DownloadFile("https://raw.githubusercontent.com/quantumwavves/KITA/master/resources/assets/365/365-full.xml", "$Path\MSOffice.xml")
      }
    }
}


function Office_Deploy {
    #Global variables
    $totalSteps=5
    $currentStep=1
    $officeName= "Microsoft Office $version"
    $DownloadUrl="://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_16731-20290.exe"
    $mirrorUrl="https://download843.mediafire.com/7soo1y8aalpghX9ZtZznTcsApj4oMJSew3iNt9uBC2z76iMHZwskdN3kK08KyRn5Z3k7SD1L1GLzUIbyRDw-5cRu1Mmj8s08Kbxk7T2j4el8NJvaj6ElQjywV05PX5dG0pXOvRmECnCor5IFGp_b4FsE5Q2bTN7g2wbd_4bYfnYfo5k/4l20xzhlo0vlwa6/officedeploymenttool_16731-20290.exe"
    #Download developement tool
    Write-Progress -Activity "Download  development deploy tool $officeName" -Status "Step $currentStep of $totalSteps" -PercentComplete (($currentStep/$totalSteps)*100)
        $HTTP_Request = [System.Net.WebRequest]::Create($DownloadUrl)
        $HTTP_Response = $HTTP_Request.GetResponse()
        $HTTP_Status = [int]$HTTP_Response.StatusCode
        if($HTTP_Status -eq "200"){
            Write-Output "[*] Status : $HTTP_Status. The download has started..."
            (New-Object System.Net.WebClient).DownloadFile($DownloadUrl, "$env:temp\officeDeploy.exe")
            Write-Output "[+] Complete download."
        }
        else{
            Write-Output "[*] Status : $HTTP_Status. Error connecting to the server, starting the download from the mirror..."
            (New-Object System.Net.WebClient).DownloadFile($mirrorUrl, "$env:temp\officeDeploy.exe")
            Write-Output "[=] Comparing hashes"
            $knowHash="613BD0952064CEF8B65335A9C50C435D5E2EDA5D7A6D0EA120806103C72BDE32"
            $srcHash = Get-FileHash $env:temp\officeDeploy.exe -Algorithm "SHA256" 
            if ($knowHash -eq $srcHash.Hash){
                Write-Output "[+] Hash status : OK"
            }else {
                Write-Error "[*] Hash status : hashes are not equal"
                Remove-Item "$env:temp\officeDeploy.exe" -Force
            }
        }
    $currentStep++
     #Status file health
     Write-Progress -Activity "Verifying files integrity $officeName" -Status "Step $currentStep of $totalSteps" -PercentComplete (($currentStep/$totalSteps)*100)
     Write-Output "[-] Verifying files integrity..."
     if (Test-Path -Path "$env:temp\officeDeploy.exe") {
         Write-Output "[+] Developement tool integrity status : OK"
     }else{
         Write-Output "[*] Developement tool integrity status : Error source not found"
     }
     if (Test-Path -Path "$env:temp\MSOffice.xml"){
        Write-Output "[+] XML file configuration integrity status : OK"
     }else{
        Write-Output "[*] XML file configuration integrity status : Error source not found"
     }
     $currentStep++
       #Unzip requiered files
    Write-Progress -Activity "Unzip files $officeName" -Status "Step $currentStep of $totalSteps" -PercentComplete (($currentStep/$totalSteps)*100)
    if (Test-Path "$env:temp\deploy" -PathType Container) {
        Remove-Item "$env:temp\deploy" -Recurse -Force
    } else {
        New-Item -Path "$env:temp" -Name "deploy" -ItemType "directory" | Out-Null
    }
    cmd.exe /c "$env:temp\officeDeploy.exe /quiet /extract:$env:temp\deploy"
    Write-Output "[*] Unzip requiered files"
    $currentStep++
    #Deploy Office 2021 LTSC version
    Write-Progress -Activity "Deploying $officeName" -Status "Step $currentStep of $totalSteps" -PercentComplete (($currentStep/$totalSteps)*100)
    Write-Output "[-] Deploy status : started, please wait"
    cmd.exe /c "$env:temp\deploy\setup.exe /configure $env:temp\MSOffice.xml"
    $currentStep++
    #Cleaning temp files
    Write-Progress -Activity "Cleaning temp files $officeName" -Status "Step $currentStep of $totalSteps" -PercentComplete (($currentStep/$totalSteps)*100)
    Remove-Item "$env:temp\deploy" -Recurse -Force
    Remove-Item "$env:temp\officeDeploy.exe" -Force
    Remove-Item "$env:temp\MSOffice.xml" -Force
    #Finished deploy
    Write-Output "[+] Deploy status : completed"
}

function postInstall {
      #Act Key management server
    $actChoise= Read-Host "[?] Do you want to activate this copy of office? (y/n) "
    while ($actChoise -ne "y" -and $actChoise -ne "n") {
            Write-Output "[*] This is not a valid value, please select a valid value"
            $actChoise= Read-Host "[*] Do you want to activate this copy of office? (y/n) "
        }
        if($actChoise -eq "y"){
            JAMSOA
            Write-Output "[*] This copy is activated"
        }
        if ($actChoise -eq "n") {Write-Output "[-] Skip activation"}
    Write-Output "[*] Finished installation"
}

function JAMSOA {
    Write-Host "#######################################"
    Write-Host "#             JAMSOA 2.0              #"
    Write-Host "#                                     #"
    Write-Host "#     0: Exit                         #"
    Write-Host "#     1: Set your own KMS server      #"
    Write-Host "#     2: Default (kms.digiboy.ir)     #"
    Write-Host "#                                     #"
    Write-Host "#######################################"
    Write-Host "                                       "
    $choise= Read-Host "[?] Select option "
    while ($choise -ne "1" -and $choise -ne "2" -and $choise -ne "0"){
        Write-Output "[-] Invalid option, please select a valid value"
        $choise = Read-Host "[?] Select option "
    }
    
    if($choise -eq "1"){$domain= Read-Host "[?] Put your kms server "}
    if($choise -eq "2"){Write-Output "[-] Default KMS Server selected"}
    if($choise -eq "0"){Write-Output "[*] The program has ended successfully"}

    $Folder = 'C:\Program Files\Microsoft Office\Office16\'
    if (Test-Path -Path $Folder) {
        Set-Location 'C:\Program Files\Microsoft Office\Office16\'
        $MSPath='C:\Program Files\Microsoft Office\Office16\'
    } else {
        Set-Location 'C:\Program Files (x86)\Microsoft Office\Office16'
        $MSPath='C:\Program Files (x86)\Microsoft Office\Office16'
    }
    $getLicenseName = cmd.exe /c "cscript ospp.vbs /dstatus" | Select-String 'License Name' 
    $desiredText = ($getLicenseName -split ',')[1] -replace 'VL_KMS_Client_AE| edition', ''
    $licenseName = $desiredText.Trim('')
    Write-Output "[*] Office version: $licenseName"
    if($licenseName -eq "Office21ProPlus2021"){
        Write-Output "[+] Installing licenses"
        $licenseFiles = Get-ChildItem -Path "..\root\Licenses16\ProPlus2021VL*.xrm-ms" -File
        foreach ($file in $licenseFiles) {
            $licensePath = $file.FullName
            & cscript.exe //nologo //B ospp.vbs /inslic:"$licensePath"
        }
        Write-Host "[-] Setting the port"
        & cscript.exe //nologo //B ospp.vbs /setprt:1688
        Write-Host "[-] Unpkey if available"
        & cscript.exe //nologo //B ospp.vbs /unpkey:6F7TH
        Write-Output "[+] Adding serial key to office"
        & cscript.exe //nologo //B ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH
        Write-Output "[+] Connect with key management server..."
        if ($choise -eq 1) {& cscript.exe //nologo //B ospp.vbs /sethst:$domain}
        if ($choise -eq 2) {& cscript.exe //nologo //B ospp.vbs /sethst:kms.digiboy.ir}
        Write-Output "[-] Activating office copy"
        & cscript.exe //nologo //B ospp.vbs /act
        Write-Output "[+] Activation completed"
        Set-Location "C:\Windows\system32"
    }
    if($licenseName -eq "Office19ProPlus2019"){
        Write-Output "[+] Installing licenses"
        $licenseFiles = Get-ChildItem -Path "..\root\Licenses16\ProPlus2019VL*.xrm-ms" -File
        foreach ($file in $licenseFiles) {
            $licensePath = $file.FullName
            & cscript.exe //nologo //B ospp.vbs /inslic:"$licensePath"
        }
        Write-Host "[-] Setting the port"
        & cscript.exe //nologo //B ospp.vbs /setprt:1688
        Write-Host "[-] Unpkey if available"
        & cscript.exe //nologo //B ospp.vbs /unpkey:6MWKP
        Write-Output "[+] Adding serial key to office"
        & cscript.exe //nologo //B ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP
        Write-Output "[+] Conect with key management server..."
        if ($choise -eq 1) {& cscript.exe //nologo //B ospp.vbs /sethst:$domain}
        if ($choise -eq 2) {& cscript.exe //nologo //B ospp.vbs /sethst:kms.digiboy.ir}
        Write-Output "[-] Activating office copy"
        & cscript.exe //nologo //B ospp.vbs /act
        Write-Output "[+] Activation completed"
        Set-Location "C:\Windows\system32"
    }
}

function RMSOA {
    $Folder = 'C:\Program Files\Microsoft Office\Office16\'
    if (Test-Path -Path $Folder) {
        Set-Location 'C:\Program Files\Microsoft Office\Office16\'
        $MSPath='C:\Program Files\Microsoft Office\Office16\'
    } else {
        Set-Location 'C:\Program Files (x86)\Microsoft Office\Office16'
        $MSPath='C:\Program Files (x86)\Microsoft Office\Office16'
    }
    $getLicenseName = cmd.exe /c "cscript ospp.vbs /dstatus" | Select-String 'License Name' 
    $desiredText = ($getLicenseName -split ',')[1] -replace 'VL_KMS_Client_AE| edition', ''
    $licenseName = $desiredText.Trim('')
    Write-Output "[*] Office version: $licenseName"

    if ($licenseName -eq "Office21ProPlus2021"){
        Set-Location "C:\Windows\system32"
        Write-Output "[-] Clear product key"
        & cscript.exe //nologo //B slmgr.vbs /cpky
        Set-Location $MSPath
        Write-Output "[-] Obtain SKU ID"
        $skuId = & cscript.exe ospp.vbs /dstatus | Select-String "SKU ID:"
        $skuId = $skuId -replace("SKU ID: ", "")
        Write-Output "[-] Your SKU ID:$skuId"
        Write-Output "[-] Unpkey microsoft office" 
        & cscript.exe //nologo //B ospp.vbs /unpkey:6F7TH
        Write-Output "[-] Reset value licenses"
        &cscript //nologo //B ospp.vbs /rearm:$skuId
        Start-Sleep -Seconds 20
        & cscript //nologo //B ospp.vbs /rearm 
        Write-Output "-[-] Clean cache host KMS"
        & cscript.exe //nologo //B ospp.vbs /ckms-domain
        Set-Location "C:\Windows\system32"
        Write-Output "[*] The key was successfully removed"
    }
    if ($licenseName -eq "Office19ProPlus2019"){
        Set-Location "C:\Windows\system32"
        Write-Output "[-] Clear product key"
        & cscript.exe //nologo //B slmgr.vbs /cpky
        Set-Location $MSPath
        Write-Output "[-] Obtain SKU ID"
        $skuId = & cscript.exe ospp.vbs /dstatus | Select-String "SKU ID:"
        $skuId = $skuId -replace("SKU ID: ", "")
        Write-Output "[-] Your SKU ID: $skuId"
        Write-Output "[-] Remove key from office"
        & cscript.exe //nologo //B ospp.vbs /unpkey:6MWKP
        Write-Output "[-] Reset value licenses"
        & cscript.exe //nologo //B ospp.vbs /rearm:$skuId
        Start-Sleep -Seconds 20 
        & cscript.exe //logo //B ospp.vbs /rearm
        Write-Output "[-] Clean cache host KMS"
        & cscript.exe //nologo //B ospp.vbs /ckms-domain 
        Write-Output "[*] The key was successfully removed"
        Set-Location "C:\Windows\system32"
    }
}

menu