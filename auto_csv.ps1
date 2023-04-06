# 設定ファイルからフォルダーパスを取得する
$config = Get-Content .\config.txt
$folder = $config.Trim()
$logFile = ".\convert.log"

# 変換関数を定義する
function ConvertTo-Excel {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateScript({ Test-Path $_ -PathType 'Leaf' })]
        [string]$CsvFilePath,
        [Parameter(Mandatory = $true)]
        [string]$ExcelFilePath
    )

    # CSVファイルを読み込む
    $csv = Import-Csv $CsvFilePath -Header '会社コード', '店舗コード', 'Jancode', 'NS販売価格'

    # Jancodeでソートされたデータを取得する
    $sorted = $csv | Sort-Object -Property Jancode

    # Excelアプリケーションオブジェクトを作成する
    $excel = New-Object -ComObject Excel.Application

    # Excelを非表示にする
    $excel.Visible = $false

    # 新しいワークブックを作成する
    $workbook = $excel.Workbooks.Add()

    # 最初のワークシートオブジェクトを取得する
    $worksheet = $workbook.Worksheets.Item(1)

    # ヘッダーを書き込む
    $worksheet.Cells.Item(1,1) = "会社コード"
    $worksheet.Cells.Item(1,2) = "店舗コード"
    $worksheet.Cells.Item(1,3) = "Jancode"
    $worksheet.Cells.Item(1,4) = "NS販売価格"

   # データを書き込む
    $row = 2
    foreach ($item in $sorted) {
        $worksheet.Cells.Item($row,1) = $item."会社コード"
        $worksheet.Cells.Item($row,2) = $item."店舗コード"
        $worksheet.Cells.Item($row,3) = $item."Jancode"
        $worksheet.Cells.Item($row,4) = $item."NS販売価格"
        $row++
    }

    # Excelファイルを保存する
    $workbook.SaveAs($ExcelFilePath)

    # ワークブックとExcelアプリケーションを閉じる
    $workbook.Close()
    $excel.Quit()

    # Excelオブジェクトを解放する
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

$logFilePath = Join-Path $env:USERPROFILE 'csv-to-excel.log'
$logMessage = "$(Get-Date) - Converted $($CsvFilePath) to $($ExcelFilePath)"
Add-Content -Path $logFilePath -Value $logMessage

# 5分ごとにフォルダーをスキャンする
while ($true) {
    Write-Host "Scanning folder: $folder"
    Get-ChildItem $folder -Filter *.csv | ForEach-Object {
        $csvPath = $_.FullName
        $excelPath = $_.FullName.Replace(".csv", ".xlsx")
        Write-Host "Converting $csvPath to $excelPath"
        try {
            ConvertTo-Excel -CsvFilePath $csvPath -ExcelFilePath $excelPath
            Remove-Item $csvPath
        } catch {
            Write-Host "Error converting $csvPath: $_"
        }
    }
    Start-Sleep -Seconds 300
}