#�Q�[���̃g�b�v�y�[�W����A���f���̃_�E�����[�h�A�h���X�𐶐�����
#$url="https://www.models-resource.com/nintendo_64/supermario64/"
#$url="https://www.models-resource.com/pc_computer/outlast/"

#�Q�[���̃g�b�v�y�[�W�̃��X�g�����
$list=get-content modre.txt

#DL�p�̃w�b�h�A�h���X
$head="https://www.models-resource.com/download/"

foreach($url in $list){
$page=Invoke-webrequest $url
$Link=$page.Links.href|Where-object {$_ -like "*/model/*"}

$G=@()

#�e���f���y�[�W�̃A�h���X���擾
foreach($k in $Link){
 $a=$k.split("/")
 $head+$a[-2]+"/" >>ddll.txt 
}

}