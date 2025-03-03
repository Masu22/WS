#エクセルの検索をPSで！
#ToDo：PSobjectで属性分け(Sheet, Cellの位置、keywordなど)をする？
#検索する単語を入力する形式にする

#参考：https://qiita.com/www-tacos/items/045dad569920cd439a9a
#文字列置換：https://qiita.com/acuo/items/a4f83d886c4b8a7fcf52 
#正規表現：https://userweb.mnet.ne.jp/nakama/

#excelオブジェクト読み込み
$EXCELAPP = New-Object -ComObject Excel.Application

#excelブックを開く、以下は変数の説明
# 1: ファイルパス
$file="C:\Users\NARUTO\Desktop\WS\PSfile\aaa.xlsx"
# 2: 0ならシート内の外部参照を更新しない
# 3: Trueなら読み取り専用で開く
# 4: テキストファイルを開く場合の区切り文字、不要なのでMissingでスキップ
# 5: パスワードがかかっている場合に試すパスワード
$password="pass"
$wb = $EXCELAPP.Workbooks.Open($file, 0, $True, [Type]::Missing, $password)

#検索キーワードを指定
$keyword="a3"

#検索結果を格納する配列
$results = @()

try{
$wb.Worksheets | Foreach-Object{
#シートの選択
$ws = $_

#シートの名前を表示
Write-Host "検索したシート：$($ws.Name)"

#セルを検索、$foundにはセルの情報が入る？
#セルをR1C1形式で表示するには以下のようにする
#Write-Host "$($found.Row),$($found.Column)"
#今回は、.addres()でA1形式で表示した。

$first = $found = $ws.Cells.Find($keyword)
$results=$found.Address()

#絶対座標の$を削除して、結果を表示
$results -replace "\$",""  #>>Search.txt

#残り全部を検索する
while ($null -ne $found) {
#次を検索
$found = $ws.Cells.FindNext($found)
$results=$found.Address()

#検索結果が$firstに戻ってきたら検索終了！！
if ($found.Address() -eq $first.Address()) {
 break
}

#結果を表示
$results -replace "\$","" #>>Search.txt
}
}
}catch{
 Write-Host "エラー発生！（文字列のないシートを検索した可能性あり）"
 break
}

#excelの終了
$wb.Quit
$EXCELAPP.Quit()
$EXCELAPP = $null
[System.GC]::Collect()
