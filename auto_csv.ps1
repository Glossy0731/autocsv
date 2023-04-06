$folder = Split-Path -Parent $MyInvocation.MyCommand.Path

$logFilePath = Join-Path -Path $folder -ChildPath ('conversion_{0:yyyyMMdd_HHmmss}.log' -f (Get-Date))

function ConvertTo-Excel {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateScript({ Test-Path $_ -PathType 'Leaf' })]
        [string]$CsvFilePath,
        [Parameter(Mandatory = $true)]
        [string]$ExcelFilePath,
        [Parameter(Mandatory = $true)]
        [string]$LogFilePath
    )

    # 导入 CSV 文件到数据表中
    $csv = Import-Csv $CsvFilePath

    # 排序数据表
    $sorted = $csv | Sort-Object -Property 会社コード, 店舗コード, Jancode

    # 创建 Excel 对象
    $excel = New-Object -ComObject Excel.Application

    # 隐藏 Excel 界面
    $excel.Visible = $false

    # 添加一个新的工作簿
    $workbook = $excel.Workbooks.Add()

    # 选择工作表
    $worksheet = $workbook.Worksheets.Item(1)

    # 写入表头
    $worksheet.Cells.Item(1,1) = "会社コード"
    $worksheet.Cells.Item(1,2) = "店舗コード"
    $worksheet.Cells.Item(1,3) = "Jancode"
    $worksheet.Cells.Item(1,4) = "NS販売価格"

    # 写入数据
    $row = 2
    foreach ($item in $sorted) {
        $worksheet.Cells.Item($row,1) = $item."会社コード"
        $worksheet.Cells.Item($row,2) = $item."店舗コード"
        $worksheet.Cells.Item($row,3) = $item."Jancode"
        $worksheet.Cells.Item($row,4) = $item."NS販売価格"
        $row++
    }

    # 保存 Excel 文件
    $workbook.SaveAs($ExcelFilePath)

    # 记录日志
    $logMessage = "{0} - {1}" -f (Get-Date), (Split-Path -Leaf $ExcelFilePath)
    Add-Content -Path $LogFilePath -Value $logMessage

    # 释放资源
    $workbook.Close()
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

# 处理所有 CSV 文件
Get-ChildItem $folder -Filter *.csv | ForEach-Object {
    $csvPath = $_.FullName
    $excelPath = Join-Path -Path $folder -ChildPath ($_.BaseName + '.xlsx')
    ConvertTo-Excel -CsvFilePath $csvPath -ExcelFilePath $excelPath -LogFilePath $logFilePath
    Remove-Item $csvPath
}