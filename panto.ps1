#cmd�Łucd ./Pantodon�v�u./panto.ps1�v���s���Ă���B

#���t����Fhttps://qiita.com/ryosuke0825/items/06eae2e99f587b5275aa
#�d���폜�Fhttps://step-learn.com/article/powershell/047-array-unique.html

#�X�V������ݒ肵�����ǁA�g��Ȃ�����
$date=Get-Date -Format "yyyy-MM-dd"

#�X�V����~�ς������X�g�̓ǂݍ���
$up=get-content C:\Users\NARUTO\Pantodon\Pantodon-Up.txt
#���p�X�Ŏw�肵�Ȃ��ƁA�t�H���_�O������s�����Ƃ��ɃG���[�ɂȂ�Ǝv����I
#�Ȃ̂ŁA$up=get-content ./Pantodon/Pantodon-Up.txt �Ȃ�

#�o�b�N�A�b�v�̍쐬
$up >>back-up-$date.txt

#���݂�HP�f�[�^��ǂݍ���
$url=Invoke-Webrequest "http://pantodon.jp/index.rb?body=about"
$A=$url.Links.href

#�f�[�^���������āA������̏d�����폜
$G=$up+$A|Select-Object -Unique

#�t�@�C���ɏ����o��
$G >Pantodon-Up.txt