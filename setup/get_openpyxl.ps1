$url = "https://files.pythonhosted.org/packages/f5/39/942a406621c1ff0de38d7e4782991b1bac046415bf54a66655c959ee66e8/openpyxl-2.6.3.tar.gz"
$output = "$PSScriptRoot\openpyxl-2.6.3.tar.gz"
$start_time = Get-Date

Invoke-WebRequest -Uri $url -OutFile $output
Write-Output "Time taken: $((Get-Date).Subtract($start_time).Seconds) second(s)"
