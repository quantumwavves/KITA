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
        Write-Output "-> Obtain SKU ID"
        $skuId = & cscript.exe ospp.vbs /dstatus | Select-String "SKU ID:"
        $skuId = $skuId -replace("SKU ID: ", "")
        Write-Output "-> Your SKU ID: $skuId"
        Write-Output "-> Remove key from office" 
        & cscript.exe //nologo //B ospp.vbs /unpkey:6F7TH
        Write-Output "-> Clean cache host KMS"
        & cscript.exe //nologo //B ospp.vbs /ckms-domain
        Write-Output "-> Reset value licenses"
        & cscript.exe ospp.vbs /rearm:$skuId
        Write-Output "-> The key was successfully removed"
        Set-Location "C:\Windows\system32"
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
        Write-Output "-> The key was successfully removed"
        Set-Location "C:\Windows\system32"
    }
}
keyRemover
