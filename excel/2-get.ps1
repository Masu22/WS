#検索結果のセルの値を取得
#row：エクセル内の一時ID
#column：項目のプロパティ(A：ID、B：項目名、C：性質などなど)
#基本的に１行で１項目なので、検索結果からヒットした行を取り出すのが目的

#参考：https://shift101.hatenablog.com/entry/2021/05/29/223218

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

#ワークシートを選択、item(何番目のシートか)で指定
$ws=$wb.Worksheets.item(1)

#セルを指定して、値を取得（配列として認識されるので、以下のように順次表示にした）
$Data = $ws.Range("A2:F2")

#取得値を順次出力
$Data | % { $_.Text }



#excelの終了
$wb.Quit
$EXCELAPP.Quit()
$EXCELAPP = $null
[System.GC]::Collect()