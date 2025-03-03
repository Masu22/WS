#ゲームのトップページから、モデルのダウンロードアドレスを生成する
#$url="https://www.models-resource.com/nintendo_64/supermario64/"
#$url="https://www.models-resource.com/pc_computer/outlast/"

#ゲームのトップページのリストを入力
$list=get-content modre.txt

#DL用のヘッドアドレス
$head="https://www.models-resource.com/download/"

foreach($url in $list){
$page=Invoke-webrequest $url
$Link=$page.Links.href|Where-object {$_ -like "*/model/*"}

$G=@()

#各モデルページのアドレスを取得
foreach($k in $Link){
 $a=$k.split("/")
 $head+$a[-2]+"/" >>ddll.txt 
}

}
