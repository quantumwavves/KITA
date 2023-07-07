#Remover Microsoft Office Activation
function keyRemover {
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
    Write-Output "-> Office version: $licenseName"

    if ($licenseName -eq "Office21ProPlus2021"){
        Set-Location "C:\Windows\system32"
        Write-Output "-> Clear product key"
        & cscript.exe //nologo //B slmgr.vbs /cpky
        Set-Location $MSPath
        Write-Output "-> Obtain SKU ID"
        $skuId = & cscript.exe ospp.vbs /dstatus | Select-String "SKU ID:"
        $skuId = $skuId -replace("SKU ID: ", "")
        Write-Output "-> Your SKU ID:$skuId"
        Write-Output "-> Unpkey microsoft office" 
        & cscript.exe //nologo //B ospp.vbs /unpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH
        Write-Output "-> Reset value licenses"
        &cscript //nologo //B ospp.vbs /rearm:$skuId
        Start-Sleep -Seconds 20
       & cscript //nologo //B ospp.vbs /rearm 
        Write-Output "-> Clean cache host KMS"
        & cscript.exe //nologo //B ospp.vbs /ckms-domain
        Set-Location "C:\Windows\system32"
        Write-Output "-> The key was successfully removed"
    }



    if ($licenseName -eq "Office19ProPlus2019"){
        Write-Output "-> Obtain SKU ID"
        $skuId = & cscript.exe ospp.vbs /dstatus | Select-String "SKU ID:"
        $skuId = $skuId -replace("SKU ID: ", "")
        Write-Output "-> Your SKU ID: $skuId"
        Write-Output "-> Remove key from office"
        & cscript.exe //nologo //B ospp.vbs /unpkey:6MWKP
        Write-Output "-> Clean cache host KMS"
        & cscript.exe //nologo //B ospp.vbs /ckms-domain
        Write-Output "-> Reset value licenses"
        & cscript.exe //nologo //B ospp.vbs /rearm:$skuId
        $skuStatus = & cscript.exe ospp.vbs /dstatus | Select-String "SKU ID:"
        if ($skuStatus -eq "") {
          & cscript.exe ospp.vbs /rearm
        }
        if ($skuId -ne "") {
          & cscript.exe ospp.vbs /rearm:$skuId
          & cscript.exe ospp.vbs /rearm
        }
        Write-Output "-> The key was successfully removed"
        Set-Location "C:\Windows\system32"
    }
}
keyRemover
