#カレントディレクトリの全ての.xlsxのA1-A1200を取り出して、outp.txtに書き出す

#https://qiita.com/23fumi/items/aa65ffa2098509337d33
#https://shift101.hatenablog.com/entry/2021/05/29/223218

#スクリプトファイルの最初に書く →これはファイル名の読み取り？今回はスキップで！
# mandatory = $trueとしている引数は必須。指定していないと聞かれる
# $visible = $true は初期値。消しておきたい場合は、実行時に -visible $falseとする
#param(
    #[parameter(mandatory=$true)]$fileName,
    #$visible = $true)


#エクセル名を指定
#$filename="1.xlsx"

# 入力元ファイル
$source = Get-Item *.xlsx

#Excelを開く
$excel = New-Object -ComObject Excel.Application
$book = $null
$excel.Visible = $false
$excel.DisplayAlerts = $false

foreach($FullPath in $source){

#ブックを開く
#$FullPath = (Get-ChildItem -Path $filename).FullName
$book = $excel.Workbooks.Open($FullPath)

#シートの選択
$currentSheet = $book.Sheets(1)

#シートの名前を取得
#$book.Sheets(1).Name

#セルの取得(rangeの場合は、配列扱い)
$RC= $currentSheet.Range("A1:A1200")

#一つずつ表示
$RC | % { $_.Text } >>outp.txt

}

#終了
[void]$book.Close()
[void]$excel.Quit()
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)