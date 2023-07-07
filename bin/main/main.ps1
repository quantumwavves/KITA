#KITA 1.0 by QuantumWavves

function kita {
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
$choise= Read-Host "`n-> Select option "
    while ($choise -ne "1" -and $choise -ne "2" -and $choise -ne "3" -and $choise -ne "4") {
      Write-Host "Please select valid option"
      $choise= Read-Host "-> Select option "
    }    
    if($choise -eq "1"){officeOpts}
    if($choise -eq "2"){jamsoa}
    if($choise -eq "3"){rmsoa}
    #if($choise -eq "4"){removeOffice} soon feature 
    if($choise -eq "4"){Write-Output "-> Exit :)"}
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
 $versionChoise = Read-Host "`n-> Select your office version"
 while ($versionChoise -ne 1 -and $versionChoise -ne 2 -and $versionChoise -ne 3) {
  Write-Output "-> Invalid option, please select valid option"
  $versionChoise = Read-Host "-> Select option"
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
 $bloatChoise = Read-Host "`n-> Select your bloat"
 while ($bloatChoise -ne 1 -and $bloatChoise -ne 2 -and $bloatChoise -ne 3 -and $bloatChoise -ne 4) {
  Write-Output "-> Invalid option, please select valid option"
  $bloatChoise = Read-Host "-> Select option"
 }
 if ($bloatChoise -eq "1") {$bloatVersion = "meta"}
 if ($bloatChoise -eq "2") {$bloatVersion = "minimal"}
 if ($bloatChoise -eq "3") {$bloatVersion = "normal"}
 if ($bloatChoise -eq "4") {$bloatVersion = "all"}
 xmlDownloader
 installer
 if ($version -eq "2019" -and $version -eq "2021") {
  postInstall
 }
 else {
  Write-Output "-> Office 365 cannot be activated by kms, use your own license" 
 }
 Write-Output "All tasks have been completed successfully exiting the program"
}

function installer {
  if ($version -eq "2019"){
      $url19 = "https://raw.githubusercontent.com/quantumwavves/KITA/master/bin/deploys/D2019ProPlus.ps1"
      Invoke-RestMethod "$url19" | Invoke-Expression
      Write-Output "-> Start installation"
    }
  if ($version -eq "2021"){
      $url21 = "https://raw.githubusercontent.com/quantumwavves/KITA/master/bin/deploys/D2021LTSC.ps1"
      Invoke-RestMethod "$url21" | Invoke-Expression
      Write-Output "-> Start installation"
    }
  if ($version -eq "365"){
      $url365 = "https://raw.githubusercontent.com/quantumwavves/KITA/master/bin/deploys/D365.ps1"
      Invoke-RestMethod "$url365" | Invoke-Expression
      Write-Output "-> Start installation"
    }
}

function xmlDownloader {
  $Path = $env:temp
  $wc = (New-Object System.Net.WebClient)
 
  if ($version -eq "2019"){
    if ($bloatVersion -eq "meta"){
      Write-Output "`n-> Donwload XML $bloatVersion configuration for deploy"
      $wc.DownloadFile("https://raw.githubusercontent.com/quantumwavves/KITA/master/resources/assets/2019/2019-meta.xml", "$Path\2019PP.xml")
      }
    if ($bloatVersion -eq "minimal"){
      Write-Output "`n-> Donwload XML $bloatVersion configuration for deploy"
      $wc.DownloadFile("https://raw.githubusercontent.com/quantumwavves/KITA/master/resources/assets/2019/2019-minimal.xml", "$Path\2019PP.xml")
      }
    if ($bloatVersion -eq "normal"){
      Write-Output "`n-> Donwload XML $bloatVersion configuration for deploy"
      $wc.DownloadFile("https://raw.githubusercontent.com/quantumwavves/KITA/master/resources/assets/2019/2019-normal.xml", "$Path\2019PP.xml")
      }
    if ($bloatVersion -eq "all"){
      Write-Output "`n-> Donwload XML $bloatVersion configuration for deploy"
      $wc.DownloadFile("https://raw.githubusercontent.com/quantumwavves/KITA/master/resources/assets/2019/2019-full.xml", "$Path\2019PP.xml")
      }
  }
  if ($version -eq "2021"){
    if ($bloatVersion -eq "meta"){
      Write-Output "`n-> Donwload XML $bloatVersion configuration for deploy"
      $wc.DownloadFile("https://raw.githubusercontent.com/quantumwavves/KITA/master/resources/assets/2021/2021-meta.xml", "$Path\2021LTSC.xml")
      }
    if ($bloatVersion -eq "minimal"){
      Write-Output "`n-> Donwload XML $bloatVersion configuration for deploy"
      $wc.DownloadFile("https://raw.githubusercontent.com/quantumwavves/KITA/master/resources/assets/2021/2021-minimal.xml", "$Path\2021LTSC.xml")
      }
    if ($bloatVersion -eq "normal"){
      Write-Output "`n-> Donwload XML $bloatVersion configuration for deploy"
      $wc.DownloadFile("https://raw.githubusercontent.com/quantumwavves/KITA/master/resources/assets/2021/2021-normal.xml", "$Path\2021LTSC.xml")
      }
    if ($bloatVersion -eq "all"){
      Write-Output "`n-> Donwload XML $bloatVersion configuration for deploy"
      $wc.DownloadFile("https://raw.githubusercontent.com/quantumwavves/KITA/master/resources/assets/2021/2021-full.xml", "$Path\2021LTSC.xml")
      }
  }
  if ($version -eq "365"){
    if ($bloatVersion -eq "meta"){
      Write-Output "`n-> Donwload XML $bloatVersion configuration for deploy"
      $wc.DownloadFile("https://raw.githubusercontent.com/quantumwavves/KITA/master/resources/assets/365/365-meta.xml", "$Path\O365.xml")
      }
    if ($bloatVersion -eq "minimal"){
      Write-Output "`n-> Donwload XML $bloatVersion configuration for deploy"
      $wc.DownloadFile("https://raw.githubusercontent.com/quantumwavves/KITA/master/resources/assets/365/365-minimal.xml", "$Path\O365.xml")
      }
    if ($bloatVersion -eq "normal"){
      Write-Output "`n-> Donwload XML $bloatVersion configuration for deploy"
      $wc.DownloadFile("https://raw.githubusercontent.com/quantumwavves/KITA/master/resources/assets/365/365-normal.xml", "$Path\O365.xml")
      }
    if ($bloatVersion -eq "all"){
      Write-Output "`n-> Donwload XML $bloatVersion configuration for deploy"
      $wc.DownloadFile("https://raw.githubusercontent.com/quantumwavves/KITA/master/resources/assets/365/365-full.xml", "$Path\O365.xml")
      }
    }
}

function postInstall {
      #Act Key management server
    $actChoise= Read-Host "-> Do you want to activate this copy of office? (y/n): "
    while ($actChoise -ne "y" -and $actChoise -ne "n") {
            Write-Output "-> This is not a valid value, please select a valid value"
            $actChoise= Read-Host "-> Do you want to activate this copy of office? (y/n): "
        }
        if($actChoise -eq "y"){
            Invoke-RestMethod "https://raw.githubusercontent.com/quantumwavves/KITA/master/bin/jamsoa/JAMSOA.ps1" | Invoke-Expression
            Write-Output "-> This copy is activated"
        }
        if ($actChoise -eq "n") {Write-Output "-> Skip activation"}
    Write-Output "-> Finished installation"
}

function jamsoa {
  $urljamsoa = "https://raw.githubusercontent.com/quantumwavves/KITA/master/bin/jamsoa/JAMSOA.ps1"
  Write-Output "-> Remember that office 365 cannot be activated with jamsoa"
  Write-Output "-> Start KMS activation"
  Invoke-RestMethod "$urljamsoa" | Invoke-Expression
  Write-Output "-> JAMSOA tasks ended, please restart your pc"
}

function rmsoa {
  $urlrmsoa = "https://raw.githubusercontent.com/quantumwavves/KITA/master/bin/jamsoa/RMSOA.ps1"
  Write-Output "-> Start remove KMS activation"
  Invoke-RestMethod "$urlrmsoa" | Invoke-Expression
  Write-Output "-> RMSOA tasks ended, please restart your pc"
}
kita
