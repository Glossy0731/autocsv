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

    # ���� CSV �ļ������ݱ���
    $csv = Import-Csv $CsvFilePath

    # �������ݱ�
    $sorted = $csv | Sort-Object -Property ���祳�`��, ���n���`��, Jancode

    # ���� Excel ����
    $excel = New-Object -ComObject Excel.Application

    # ���� Excel ����
    $excel.Visible = $false

    # ���һ���µĹ�����
    $workbook = $excel.Workbooks.Add()

    # ѡ������
    $worksheet = $workbook.Worksheets.Item(1)

    # д���ͷ
    $worksheet.Cells.Item(1,1) = "���祳�`��"
    $worksheet.Cells.Item(1,2) = "���n���`��"
    $worksheet.Cells.Item(1,3) = "Jancode"
    $worksheet.Cells.Item(1,4) = "NS؜�Ӂ���"

    # д������
    $row = 2
    foreach ($item in $sorted) {
        $worksheet.Cells.Item($row,1) = $item."���祳�`��"
        $worksheet.Cells.Item($row,2) = $item."���n���`��"
        $worksheet.Cells.Item($row,3) = $item."Jancode"
        $worksheet.Cells.Item($row,4) = $item."NS؜�Ӂ���"
        $row++
    }

    # ���� Excel �ļ�
    $workbook.SaveAs($ExcelFilePath)

    # ��¼��־
    $logMessage = "{0} - {1}" -f (Get-Date), (Split-Path -Leaf $ExcelFilePath)
    Add-Content -Path $LogFilePath -Value $logMessage

    # �ͷ���Դ
    $workbook.Close()
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

# �������� CSV �ļ�
Get-ChildItem $folder -Filter *.csv | ForEach-Object {
    $csvPath = $_.FullName
    $excelPath = Join-Path -Path $folder -ChildPath ($_.BaseName + '.xlsx')
    ConvertTo-Excel -CsvFilePath $csvPath -ExcelFilePath $excelPath -LogFilePath $logFilePath
    Remove-Item $csvPath
}