#Chrome�ŃA�h���X���J���āA�����_�E�����[�h

# URL���X�g�̃e�L�X�g�t�@�C���̃p�X
$urlsFile = "urls.txt"

# Chrome�̃p�X�i�f�t�H���g�̃C���X�g�[����j
$chromePath = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"

# �G���[���O�t�@�C���̃p�X
$errorLog = "error_log.txt"

# �G���[���O���N���A�i�O��̎��s���ʂ������j
Clear-Content -Path $errorLog -ErrorAction SilentlyContinue

# URL���X�g��1�s���ǂݍ���ŏ���
Get-Content $urlsFile | ForEach-Object {
    $url = $_.Trim()
    if ($url -ne "") {  # ��s���X�L�b�v
        Write-Host "Opening: $url"
        try {
            Start-Process -FilePath $chromePath -ArgumentList $url -ErrorAction Stop
            Start-Sleep -Seconds 15  # �_�E�����[�h�ҋ@���ԁi�K�v�Ȃ璲���j
        } catch {
            $errorMessage = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Error opening $url : $_"
            Write-Host $errorMessage
            Add-Content -Path $errorLog -Value $errorMessage
        }
    }
}
Write-Host "All URLs processed!"

