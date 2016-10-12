param( 
  [Parameter( Mandatory=$true)]	
  [string]$csv
)

$i = 0
$skuFlag = 0
$loopCounter = 0
$out = @()
$userList = Import-Csv $csv
$userCount = $userList.Count
$tenantName = (Get-MsolCompanyInformation).DisplayName
$skus = Get-MsolAccountSku
$skuList =  $skus | Out-String
$skuMessage = "Below is a list of all the availble SKUs in the tenant for $($tenantName):`n $skuList"
$resultsPath = $env:USERPROFILE + "\Desktop\results_" + ($csv.Split('\')[-1])


if ( !$userCount ) {
  Write-Host "`n`nThe specified CSV is emtpy or contains only 1 user; please use the GUI to make desired changes." -ForegroundColor Red
  Start-Sleep -Seconds 5
  cls
  Exit
}

function disableService {
  $disable = Read-Host "`nDo you want to disable any services when assiging the license $addLic to users? (yes/no)"
  if ( $disable -like "yes" ) {
    $skuService = ( $skus | ? { $_.AccountSKUID -match $addLic } ).ServiceStatus |`
    ? {$_.ProvisioningStatus -eq "Success"} | Select -ExpandProperty ServicePlan | Select ServiceName
    $skuServiceList = $skuService | Out-String
    cls
    Write-Host "`nBelow is a list of all services included in the license $addLic :`n $skuServiceList" -ForegroundColor Yellow
    $disabledServices = Read-Host "Please enter the service name(s) you wish NOT to enable when assigning the new license,`nleave blank for none, use the following format; $($skuService[0].ServiceName), $($skuService[$skuService.GetUpperBound(0)].ServiceName)"
    cls

    if ( $disabledServices -eq $null -or $disabledServices -eq "") {
      $script:licOpt = $null
      Read-Host "`nYou did not specify any services, all services will be enabled when assiging the new license, press any key to continue"
    } else {
      $script:licOpt = New-MsolLicenseOptions -AccountSkuId $addLic -DisabledPlans $disabledServices
      $script:licoptlist = ($licOpt.DisabledServicePlans.Replace(" ","").Split(",") | Out-String).Trim()
      $script:licoptlist = $licoptlist -replace "(?m)^", "`t"
    }

  } elseif ( $disable -like "no" ){
    $script:licOpt = $null
  } else { Exit }

}

function setAlpha2 {
Write-Host "`nBefore performing any action we will verify all users have an assigned country usage location set.
An alpha-2 code of a country of your choice is required, in order to assign to any user with no usage location.

Example alpha-2 country codes; US, GB, CH, MX. For a full list of countries and their alpha-2 codes please visit: 
http://www.nationsonline.org/oneworld/country_code_list.htm" -ForegroundColor Yellow

$alpha2code = Read-Host "`nPlease enter the desired alpha-2 code"
$script:alpha2code = $alpha2code.Trim()
cls
}

function findUser { 

  try { $user = Get-MsolUser -UserPrincipalName $a.UserPrincipalName -ErrorAction SilentlyContinue } Catch {}
  if ( !$user ) {
    $script:action = "None"
    $script:result = "User not found in Office365"
  }
  return $user
}

function checkUsageLocation {

  if ( $targetUser.UsageLocation -eq $null ) {
    Set-MsolUser -UserPrincipalName $targetUser.UserPrincipalName -UsageLocation $alpha2code
  }

}

function progress {
  Write-Progress -Activity "Running Licesing Operation" -Status "Processing user $($a.UserPrincipalName)" -Id 1 -PercentComplete ($i / $userCount * 100)
}

function addLicense {
  $script:action = "Add"
  if ( $targetUser ) {

    if ( $targetUser.Licenses.AccountSkuId -notcontains $addLic ) {
    
      if ( $licOpt -eq $null ) {
        try { Set-MsolUserLicense -UserPrincipalName $targetUser.UserPrincipalName -AddLicenses $addLic `
          -ErrorAction Stop } catch { $addErr = $_ }
      } else {
        try { Set-MsolUserLicense -UserPrincipalName $targetUser.UserPrincipalName -AddLicenses $addLic `
           -LicenseOptions $licOpt -ErrorAction Stop } catch { $addErr = $_ }       
      }
    
      if ( $addErr ) {
        $script:result = $addErr
      } else {
        $script:result = "Succesfully added license $addLic"
      }

    } else {
      $script:action = "None"
      $script:result = "No need to add, license $addLic already assigned"
    }

  }
  
}

function removeLicense {
  $script:action = "Remove"
  if ( $targetUser ) {
  
    if ( $targetUser.Licenses.AccountSkuId -contains $removeLic ) {
      try { Set-MsolUserLicense -UserPrincipalName $targetUser.UserPrincipalName -RemoveLicenses $removeLic -ErrorAction Stop } catch { $remErr = $_ }
      if ( $remErr ) {
        $script:result = $remErr
      } else {
       $script:result = "Succesfully removed license $removeLic"
      }
   
    } else {
      $script:action = "None"
      $script:result = "No need to remove, license $removeLic not assigned"
    }

  }

}

function replaceLicense {
  $script:action = "Replace"
  if ( $targetUser.Licenses.AccountSkuId -contains $removeLic -and $targetUser.Licenses.AccountSkuId -notcontains $addLic ) {
          
    if ( $licOpt -eq $null ) {
      try { Set-MsolUserLicense -UserPrincipalName $targetUser.UserPrincipalName -AddLicenses $addLic `
        -RemoveLicenses $removeLic -ErrorAction Stop } catch { $replaceErr = $_ }
    } else {
      try { Set-MsolUserLicense -UserPrincipalName $targetUser.UserPrincipalName -AddLicenses $addLic `
      -RemoveLicenses $removeLic -LicenseOptions $licOpt -ErrorAction Stop } catch { $replaceErr = $_ }       
    }

    if ( $replaceErr ) {
      $script:result = $replaceErr
    } else {
      $script:result = "Succesfully replaced license $removeLic with $addLic"
    }

  } elseif ( $targetUser.Licenses.AccountSkuId -notcontains $removeLic -and $targetUser.Licenses.AccountSkuId -contains $addLic ) {
    $script:action = "None"
    $script:result = "No need to replace as licensing meets desired state"
  } elseif ( $targetUser.Licenses.AccountSkuId -contains $removeLic -and $targetUser.Licenses.AccountSkuId -contains $addLic ) {
    removeLicense
  } elseif ( $targetUser.Licenses.AccountSkuId -notcontains $removeLic -and $targetUser.Licenses.AccountSkuId -notcontains $addLic ){
    addLicense
  }

}

Write-Host "`nThis script will allow you to add/remove/replace a license or enable/disable services
for the current license on all users in the specified CSV.`n" -ForegroundColor Yellow
Write-Host -Object $skuMessage -ForegroundColor Yellow
$targetLic = Read-Host "Please enter the SKU ID of the license you wish to add/remove/replace or enable/disable services for"

Do {
  if ( !$skus.Accountskuid.Contains($targetLic) ) {
    cls
    Write-Host "`n`nYou did not enter a vaild SKU ID or it is blank." -ForegroundColor Red
    Write-Host -Object $skuMessage -ForegroundColor Yellow
    $targetLic = Read-Host "Please enter the SKU ID of the license you wish to add/remove/replace or enable/disable services for"
    $loopCounter++
  } else {
    $skuFlag = 1
  }
} Until ( $skuFlag -eq 1 -or $loopCounter -eq 3)
if ( $skuFlag -ne 1) { Exit }

cls
Write-Host "`n"
Read-Host "Press any key to continue and select the licensing operation to perform on $targetLic"

$action = "Add", "Replace", "Remove", "Enable", "Disable" | Out-GridView -PassThru -Title "Choose license operation"

cls

if ( $action -match "Enable" -or $action -match "Disable" ) {
  Write-Host "`nYou chose to disable or enable a service(s) for all users specified that have the license $targetLic assigned." -ForegroundColor Yellow
} elseif ( $action -match "Add" ) {
  $addLic = $targetLic
  setAlpha2
  disableService
  cls
  Write-Host "`nThe following action will be performed on the $userCount users in the provided CSV:" -ForegroundColor Green
  Write-Host "`n*********************************************************************************" -ForegroundColor Green
  Write-Host "`n1. The license $addLic will be assigned." -ForegroundColor Green
  if ( $licOpt -ne $null ) {
    Write-Host "`n1a. The following services will not be enabled when assiging $addLic ;`n`n$($licOptList) " -ForegroundColor Green
  }

  Write-Host "`n*********************************************************************************" -ForegroundColor Green
  $run = Read-Host "`nDo you want to proceed? (Yes/No)?"
  if ( $run -match "Yes" ) {
    
    foreach ( $a in $userList ) {
      $i++
      progress
      $targetUser = findUser
      if ( $targetUser -ne $null ) {
        checkUsageLocation
        addLicense
      }

      $obj = [pscustomobject]@{
        Username = $a.UserPrincipalName
        Action = $action
        Result = $result
      }
      
      $out += $obj
    }

  } elseif ( $run -match "No" ) {
    Write-Host "`nYou selected to not commit the above operation, no changes will be made and the script will now quit." -ForegroundColor Red
    Start-Sleep -Seconds 5
    cls
    Exit
  }

} elseif ( $action -match "Replace" ) {
  $removeLic = $targetLic
  setAlpha2
  Write-Host "You chose to remove $removeLic from all users specified and assign a new license.`n $skuMessage" -ForegroundColor Yellow
  $addLic = Read-Host "Please enter the SKU ID of the license you want to assign to all specified users"
  $skuFlag = 0
  $loopCounter = 0
  Do {
    if ( !$skus.Accountskuid.Contains($addLic) ) {
      cls
      Write-Host "`n`nYou did not enter a vaild SKU ID or it is blank." -ForegroundColor Red
      Write-Host -Object $skuMessage -ForegroundColor Yellow
      $addLic = Read-Host "Please enter the SKU ID of the license you want to assign to all specified users"
      $loopCounter++
    } else {
      $skuFlag = 1
    }
  } Until ( $skuFlag -eq 1 -or $loopCounter -eq 3)
  
  if ( $skuFlag -ne 1) { Exit }
  disableService
  cls
  Write-Host "`nThe following action will be performed on the $userCount users in the provided CSV:" -ForegroundColor Green
  Write-Host "`n*********************************************************************************" -ForegroundColor Green
  Write-Host "`n1. The license $removeLic will be REMOVED and REPLACED with $addLic." -ForegroundColor Green
  if ( $licOpt -ne $null ) {
    Write-Host "`n1a. The following services will not be enabled when assiging $addLic ;`n`n$($licOptList) " -ForegroundColor Green
  }
  Write-Host "`n*********************************************************************************" -ForegroundColor Green
  $run = Read-Host "`nDo you want to proceed? (Yes/No)?"
  if ( $run -match "Yes" ) {
    
    foreach ( $a in $userList ) {
      $i++
      progress
      $targetUser = findUser
      if ( $targetUser -ne $null ) {
        checkUsageLocation
        replaceLicense
      }

      $obj = [pscustomobject]@{
        Username = $a.UserPrincipalName
        Action = $action
        Result = $result
      }
      
      $out += $obj
    }

  } elseif ( $run -match "No" ) {
    Write-Host "`nYou selected to not commit the above operation, no changes will be made and the script will now quit." -ForegroundColor Red
    Start-Sleep -Seconds 5
    cls
    Exit
  } else { 
    Write-Host "`nInvalid option entered, the script will now exit." -ForegroundColor Red
    Start-Sleep -Seconds 5
    cls
    Exit
  }
  
} elseif ( $action -match "Remove" ) {
  cls
  $targetLic = $removeLic
  Write-Host "`nThe following action will be performed on the $userCount users in the provided CSV:" -ForegroundColor Green
  Write-Host "`n*********************************************************************************" -ForegroundColor Green
  Write-Host "`n1. The license $removeLic will be removed." -ForegroundColor Green
  Write-Host "`n*********************************************************************************" -ForegroundColor Green
  $run = Read-Host "`nDo you want to proceed? (Yes/No)?"
  if ( $run -match "Yes" ) {
    
    foreach ( $a in $userList ) {
      $i++
      progress
      $targetUser = findUser
      if ( $targetUser -ne $null ) {
        removeLicense
      }

      $obj = [pscustomobject]@{
        Username = $a.UserPrincipalName
        Action = $action
        Result = $result
      }
      
      $out += $obj
    }
  
  } elseif ( $run -match "No" ) {
    Write-Host "`nYou selected to not commit the above operation, no changes will be made and the script will now quit." -ForegroundColor Red
    Start-Sleep -Seconds 5
    cls
    Exit
  } else { 
    Write-Host "`nInvalid option entered, the script will now exit." -ForegroundColor Red
    Start-Sleep -Seconds 5
    cls
    Exit
  }

} else {
  Write-Host "You did not select a licensing action in the pop-up menu, existing script" -ForegroundColor Red
  Exit
}

$out | Export-CSV -NoTypeInformation -Path $resultsPath
Write-Host "`nThe log file for this operation is saved here:`n $resultsPath" -ForegroundColor Green



#List all SKUs for tenant
#Select SKU to target
#Ask to add/replace/remove sku or enable/disable services for target SKU
  #if add
    #check for user
    #ask to disable services for target SKU
      #if disable, list all services for destination SKU
      #select services
    #check if target SKU is present
      #if present, do nothing
      #if not, add target SKU
  #if enable/disable services, list services for targeted SKU
    #Select services
    #check for user
    #check if target SKU is present
      #if present, enable/disable services for target SKU
      #if not, add target SKU and enable/disable services
  #if replace, select destination SKU
    #ask to disable services for destination SKU
      #if disable, list all services for destination SKU
      #select services
    #check for user
    #check if dest SKU is present
      #if not present check if target SKU is not present
        #if not present just add dest SKU 
        #if present replace targeted SKU with destination SKU
      #if present check if target SKU is present
        #if present, just remove target
        #if not, do nothing
  #if remove
    #check for user
    #check if target SKU is present
      #if present, remove target SKU
      #if not, do nothing
#output results