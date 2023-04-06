# �O���ե����뤫��ե�����`�ѥ���ȡ�ä���
$config = Get-Content .\config.txt
$folder = $config.Trim()
$logFile = ".\convert.log"

# ��Q�v�����x����
function ConvertTo-Excel {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateScript({ Test-Path $_ -PathType 'Leaf' })]
        [string]$CsvFilePath,
        [Parameter(Mandatory = $true)]
        [string]$ExcelFilePath
    )

    # CSV�ե�������i���z��
    $csv = Import-Csv $CsvFilePath -Header '���祳�`��', '���n���`��', 'Jancode', 'NS؜�Ӂ���'

    # Jancode�ǥ��`�Ȥ��줿�ǩ`����ȡ�ä���
    $sorted = $csv | Sort-Object -Property Jancode

    # Excel���ץꥱ�`����󥪥֥������Ȥ����ɤ���
    $excel = New-Object -ComObject Excel.Application

    # Excel��Ǳ�ʾ�ˤ���
    $excel.Visible = $false

    # �¤�����`���֥å������ɤ���
    $workbook = $excel.Workbooks.Add()

    # ����Υ�`�����`�ȥ��֥������Ȥ�ȡ�ä���
    $worksheet = $workbook.Worksheets.Item(1)

    # �إå��`������z��
    $worksheet.Cells.Item(1,1) = "���祳�`��"
    $worksheet.Cells.Item(1,2) = "���n���`��"
    $worksheet.Cells.Item(1,3) = "Jancode"
    $worksheet.Cells.Item(1,4) = "NS؜�Ӂ���"

   # �ǩ`��������z��
    $row = 2
    foreach ($item in $sorted) {
        $worksheet.Cells.Item($row,1) = $item."���祳�`��"
        $worksheet.Cells.Item($row,2) = $item."���n���`��"
        $worksheet.Cells.Item($row,3) = $item."Jancode"
        $worksheet.Cells.Item($row,4) = $item."NS؜�Ӂ���"
        $row++
    }

    # Excel�ե�����򱣴椹��
    $workbook.SaveAs($ExcelFilePath)

    # ��`���֥å���Excel���ץꥱ�`�������]����
    $workbook.Close()
    $excel.Quit()

    # Excel���֥������Ȥ��Ť���
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

$logFilePath = Join-Path $env:USERPROFILE 'csv-to-excel.log'
$logMessage = "$(Get-Date) - Converted $($CsvFilePath) to $($ExcelFilePath)"
Add-Content -Path $logFilePath -Value $logMessage

# 5�֤��Ȥ˥ե�����`�򥹥���󤹤�
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