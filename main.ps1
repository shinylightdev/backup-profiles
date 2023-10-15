
# Modify this - WHERE TO SAVE
$targetSaveLocation = 'C:\temp\save'
$pathChromeUserData = "$env:LOCALAPPDATA\Google\Chrome\User Data"
$profiles = @{}
                          

$json = Get-Content -Raw "$pathChromeUserData\Local State" | ConvertFrom-Json
$jsonFragment = $json.profile

foreach ($obj in $jsonFragment) {   
  $options = $obj.info_cache
  $propertyNames = $options.psobject.Properties.Name  
  
  foreach ($name in $propertyNames) {      
    $profiles["$name"] = $options.$name.shortcut_name      
  }
}

foreach ($p in $profiles.keys) {    
  $newDirName = "${p} ($($profiles.$p))"  
  mkdir "$targetSaveLocation\$newDirName"
  Copy-Item "$pathChromeUserData\${p}\Bookmarks" "$targetSaveLocation\$newDirName"  
}



