# Pre-filled options
$preFilledUserName = ""
$preFilledWildcard = "DWR45*"
$preFilledOU = "OU=47678"
$preFilledPassword = ""
$skipConfirmation = $false

# Run the primary script with pre-filled options
.\PrimaryScript.ps1 -UserName $preFilledUserName -Wildcard $preFilledWildcard -OU $preFilledOU -Password $preFilledPassword -SkipConfirmation:$skipConfirmation
