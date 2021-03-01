param(
    # フォルダパス取得
    # 印刷開始、終了ページ数の指定
    [parameter(mandatory)][string]$filepath,
    [parameter(mandatory)][int]$startpage,
    [parameter(mandatory)][int]$endpage

)

# フォルダ内のExcelファイル取得
$fileitems = Get-ChildItem $filepath -Recurse -File -Include *.xls,*.xlsx

# Excelファイルループ
foreach($fileitem in $fileitems)
{
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $wb = $excel.Workbooks.Open($fileitem.FullName)

        # 保存PDF名取得
        $pdfpath = $fileitem.DirectoryName + "\" + $fileitem.BaseName + ".pdf"

        # PDF出力 ファイルパス、品質通常、プロパティ含む、印刷範囲有効、印刷開始、印刷終了
        $wb.ExportAsFixedFormat([Microsoft.Office.Interop.Excel.XlFixedFormatType]::xlTypePDF, $pdfpath,0,'True','False',$startpage,$endpage)

        $wb.Close()

        $excel.Quit()
    }
    finally {
        # オブジェクト解放
        $sheet, $wb, $excel | ForEach-Object {
            if ($_ -ne $null) {
                [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($_)
            }
        }
    }
}