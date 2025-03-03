#�J�����g�f�B���N�g���̑S�Ă�.xlsx��A1-A1200�����o���āAoutp.txt�ɏ����o��

#https://qiita.com/23fumi/items/aa65ffa2098509337d33
#https://shift101.hatenablog.com/entry/2021/05/29/223218

#�X�N���v�g�t�@�C���̍ŏ��ɏ��� ������̓t�@�C�����̓ǂݎ��H����̓X�L�b�v�ŁI
# mandatory = $true�Ƃ��Ă�������͕K�{�B�w�肵�Ă��Ȃ��ƕ������
# $visible = $true �͏����l�B�����Ă��������ꍇ�́A���s���� -visible $false�Ƃ���
#param(
    #[parameter(mandatory=$true)]$fileName,
    #$visible = $true)


#�G�N�Z�������w��
#$filename="1.xlsx"

# ���͌��t�@�C��
$source = Get-Item *.xlsx

#Excel���J��
$excel = New-Object -ComObject Excel.Application
$book = $null
$excel.Visible = $false
$excel.DisplayAlerts = $false

foreach($FullPath in $source){

#�u�b�N���J��
#$FullPath = (Get-ChildItem -Path $filename).FullName
$book = $excel.Workbooks.Open($FullPath)

#�V�[�g�̑I��
$currentSheet = $book.Sheets(1)

#�V�[�g�̖��O���擾
#$book.Sheets(1).Name

#�Z���̎擾(range�̏ꍇ�́A�z�񈵂�)
$RC= $currentSheet.Range("A1:A1200")

#����\��
$RC | % { $_.Text } >>outp.txt

}

#�I��
[void]$book.Close()
[void]$excel.Quit()
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)