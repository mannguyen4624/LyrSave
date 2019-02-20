# author: Webster
# gets info from Lyrania API to output changes to files
function lyraniaInfo() {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    $user = "Webster"
    $url = "https://lyrania.co.uk/api/accounts/public_profile.php?search=$user"
    $result = (Invoke-WebRequest -Uri $url).Content.Trim()

    $json = ConvertFrom-Json($result)
    csvWrite($json)
}

# author: Webster
# comments by @nitoplol
# writes all public player info into a csv file
function csvWrite($jsonInfo) {

    # where to put csv file
    $fileName = "lyraniaData.csv"
    Set-Location ~\Documents

    # if csv file does not already exist
    if (-NOT (Test-Path $fileName)) {
        #csv does not exist
        $csvStuff = "date, memberid, name, level, quests_completed, kills, base_stats, earned_dp, guild_name"
        New-Item -Name "$fileName" -ItemType "file"
        $csvStuff | Out-File $fileName -Append
    }

    # get each item from public profile
    $date = (Get-Date -UFormat "%Y%m%d")
    $memberid = $jsonInfo.id
    $name = $jsonInfo.name
    $level = $jsonInfo.level
    $quests = $jsonInfo.quests_complete
    $kills = $jsonInfo.kills
    $base_stats = $jsonInfo.base_stats
    $earned_dp = $jsonInfo.earned_dp
    $guild = $jsonInfo.guild_name

    # set all to string
    $output = "$date, $memberid, $name, $level, $quests, $kills, $base_stats, $earned_dp, $guild"

    # write to csv file
    $output | Out-File $fileName -Append
}

# author: @nitoplol
# gets personal info from Lyrania API and writes to excel spreadsheet
function updateMoney() {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    $user = "itzu"
    $url = "https://lyrania.co.uk/api/accounts/public_profile.php?search=$user"

    # try to get values
    try {$result = (Invoke-WebRequest -Uri $url -UseBasicParsing).Content.Trim()}
    # TODO: what error is thrown when certificate is not renewed?
    catch {break}

    $json = ConvertFrom-Json($result)
    excelWrite($json)
}

# author: @nitoplol
# gets personal information and writes to an excel file
function excelWrite($jsonInfo) {
    # set path of file and sheetname
    $path = "C:\Users\Sytec\Documents\LyrSave\lyr.xlsx"
    $sn = "Sheet1"

    # declare which personal info needed
    $money = "9"
    $level = $jsonInfo.level
    $kills = $jsonInfo.kills
    $base_stats = $jsonInfo.base_stats

    # open excel in powershell
    $excel = New-Object -ComObject Excel.Application

    # don't open up excel in windows
    $excel.Visible = $false

    # get the workbook needed
    $wb = $excel.Workbooks.Open($path)

    # get the worksheet needed
    $ws = $wb.sheets.item($sn)

    # set specific cells to retrieved values
    #$moneyCell = "H2"
    #$ws.Range($moneyCell).Value = $money

    $dateCell = "E1"
    $lastDateNum = $ws.Range($dateCell).Value2
    $lastDate = [DateTime]::FromOADate($lastDateNum)

    $date = (Get-Date -UFormat "%m/%d")
    if (-NOT($lastDate -eq $date)) {
        $counter = 2
        do {
            $counter++
        } until ($ws.Range("E$counter").Text -eq "")
        $ws.Range("E3", "E$counter").clearcontents()
        $ws.Range("F3", "F$counter").clearcontents()
        $ws.Range("H3", "H$counter").clearcontents()
        $ws.Range("J3", "J$counter").clearcontents()
    }
    $ws.Range($dateCell).Value = $date

    $counter = 2
    do {
        $counter++
    } until ($ws.Range("E$counter").Text -eq "")
    $ws.Range("E$counter").Value = (Get-Date -UFormat "%T")
    $ws.Range("F$counter").Value = $level
    $ws.Range("H$counter").Value = $kills
    $ws.Range("J$counter").Value = $base_stats

    # close, quit, and remove to allow access after script runs
    $wb.Close($true)
    $excel.Quit()
    Remove-Variable -Name excel
}

updateMoney