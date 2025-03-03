#1つのエクセルファイルにまとめる！
#https://obasan.net/2024/01/06/%E8%A4%87%E6%95%B0%E3%81%AEexcel%E3%83%95%E3%82%A1%E3%82%A4%E3%83%AB%E3%81%AE%E3%82%B7%E3%83%BC%E3%83%88%E3%82%921%E3%81%A4%E3%81%AEexcel%E3%83%96%E3%83%83%E3%82%AF%E3%81%B8%E3%82%B3%E3%83%94%E3%83%BC/

$excel = New-Object -ComObject Excel.Application
$book = $null
 
$excel.Visible = $false
$excel.DisplayAlerts = $false
 
# 出力先ファイル
$destFilePath = (Convert-Path .) + "\AllSheets.xlsx"
$destBook = $excel.Workbooks.add()
 
# 入力元ファイル
$sourceFiles = Get-Item *.xlsx
 
foreach($item in $sourceFiles){
    $sourceBook = $excel.Workbooks.Open($item)
 
    foreach($s in $sourceBook.sheets){
        $sourceBook.Worksheets.item($s.Name).copy([System.Reflection.Missing]::Value,$destBook.Worksheets.item($destBook.worksheets.count))
    }
 
    [void]$sourceBook.Close($false)
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($sourceBook)
 
}
 
# AllSheetsの初期シートは削除
$destBook.Worksheets.item(1).delete()
 
[void]$destBook.SaveAs($destFilePath)
 
[void]$destBook.Close($false)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($destBook)
 
[void]$excel.Quit()
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
 
Pause