#Office 365 Deploy
function O365 {
    #Global variables
    $totalSteps=5
    $currentStep=1
    $officeVersion="365"
    $DownloadUrl="https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_16626-20148.exe"
    $mirrorUrl="https://media.githubusercontent.com/media/quantumwavves/KITA/master/resources/executables/officedeploymenttool_16626-20148.exe"
    #Download developement tool
    Write-Progress -Activity "Download  development deploy tool $officeVersion" -Status "Step $currentStep of $totalSteps" -PercentComplete (($currentStep/$totalSteps)*100)
        $HTTP_Request = [System.Net.WebRequest]::Create($DownloadUrl)
        $HTTP_Response = $HTTP_Request.GetResponse()
        $HTTP_Status = [int]$HTTP_Response.StatusCode
        if($HTTP_Status -eq "200"){
            Write-Output "-> Status : $HTTP_Status. The download has started..."
            (New-Object System.Net.WebClient).DownloadFile($DownloadUrl, "$env:temp\officeDeploy.exe")
            Write-Output "-> Complete download."
        }
        else{
            Write-Output "-> Status : $HTTP_Status. Error connecting to the server, starting the download from the mirror..."
            (New-Object System.Net.WebClient).DownloadFile($mirrorUrl, "$env:temp\officeDeploy.exe")
            Write-Output "-> Comparing hashes"
            $knowHash="f2df38c4764721a8c0b8a67af43847ae149e4fe075eadf343344afc9a79f6675"
            $srcHash = Get-FileHash $env:temp\officeDeploy.exe -Algorithm "SHA256" 
            if ($knowHash -eq $srcHash.Hash){
                Write-Output "-> Hash status : OK"
            }else {
                Write-Error "-> Hash status : hashes are not equal"
                Remove-Item "$env:temp\officeDeploy.exe" -Force
            }
        }
    $currentStep++
     #Status file health
     Write-Progress -Activity "Verifying files integrity $officeVersion" -Status "Step $currentStep of $totalSteps" -PercentComplete (($currentStep/$totalSteps)*100)
     Write-Output "-> Verifying files integrity..."
     if (Test-Path -Path "$env:temp\officeDeploy.exe") {
         Write-Output "-> Developement tool integrity status : OK"
     }else{
         Write-Output "-> Developement tool integrity status : Error source not found"
     }
     if (Test-Path -Path "$env:temp\O365.xml"){
        Write-Output "-> XML file configuration integrity status : OK"
     }else{
        Write-Output "-> XML file configuration integrity status : Error source not found"
     }
     $currentStep++
       #Unzip requiered files
    Write-Progress -Activity "Unzip files $officeVersion" -Status "Step $currentStep of $totalSteps" -PercentComplete (($currentStep/$totalSteps)*100)
    if (Test-Path "$env:temp\deploy" -PathType Container) {
        Remove-Item "$env:temp\deploy" -Recurse -Force
    } else {
        New-Item -Path "$env:temp" -Name "deploy" -ItemType "directory" | Out-Null
    }
    cmd.exe /c "$env:temp\officeDeploy.exe /quiet /extract:$env:temp\deploy"
    Write-Output "-> Unzip requiered files"
    $currentStep++
    #Deploy Office 365 LTSC version
    Write-Progress -Activity "Deploying office $officeVersion" -Status "Step $currentStep of $totalSteps" -PercentComplete (($currentStep/$totalSteps)*100)
    Write-Output "-> Deploy status : started, please wait"
    cmd.exe /c "$env:temp\deploy\setup.exe /configure $env:temp\O365.xml" 
    $currentStep++
    #Cleaning temp files
    Write-Progress -Activity "Cleaning temp files $officeVersion" -Status "Step $currentStep of $totalSteps" -PercentComplete (($currentStep/$totalSteps)*100)
    Remove-Item "$env:temp\deploy" -Recurse -Force
    Remove-Item "$env:temp\officeDeploy.exe" -Force
    Remove-Item "$env:temp\O365.xml" -Force
    #Finished deploy
    Write-Output "-> Deploy status : completed"
}
O365
