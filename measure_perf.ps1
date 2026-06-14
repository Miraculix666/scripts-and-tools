# Create dummy data
$dummyUsers = 1..1000 | ForEach-Object {
    [PSCustomObject]@{
        SamAccountName = "User$_"
    }
}

$outputFile1 = "out1.txt"
$outputFile2 = "out2.txt"

if (Test-Path $outputFile1) { Remove-Item $outputFile1 }
if (Test-Path $outputFile2) { Remove-Item $outputFile2 }

$measure1 = Measure-Command {
    $dummyUsers | ForEach-Object { $_.SamAccountName | Out-File -Append -FilePath $outputFile1 }
}

$measure2 = Measure-Command {
    $dummyUsers.SamAccountName | Out-File -Append -FilePath $outputFile2
}

$measure3 = Measure-Command {
    $dummyUsers | Select-Object -ExpandProperty SamAccountName | Out-File -Append -FilePath $outputFile2
}

Write-Host "Baseline: $($measure1.TotalMilliseconds) ms"
Write-Host "Optimized (.SamAccountName): $($measure2.TotalMilliseconds) ms"
Write-Host "Optimized (Select-Object): $($measure3.TotalMilliseconds) ms"
