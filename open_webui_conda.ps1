#requires -RunAsAdministrator
Add-Type -Name Win32 -Namespace Console -MemberDefinition @'
  [DllImport("kernel32.dll")]
  public static extern IntPtr GetConsoleWindow();
  [DllImport("user32.dll")]
  public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
'@

$consolePtr = [Console.Win32]::GetConsoleWindow()
if ($consolePtr -ne [IntPtr]::Zero) {
    # 0은 SW_HIDE: 콘솔 창을 숨김
    [Console.Win32]::ShowWindow($consolePtr, 0)
}
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# 바로가기 옵션 전역 변수
$global:selectedLanguage = "한국어"
$global:createDesktopShortcut = $true

# 다국어 리소스
$global:translations = @{
    "한국어" = @{
        "separator" = "==================================================="
        "openwebui_installing" = "Open WebUI 설치 중..."

        # 초기 설정 폼
        "setup_title" = "Open WebUI 설치 마법사"
        "setup_description" = "Open WebUI를 설치하기 위한 기본 설정을 선택해 주세요."
        "language_group" = "언어 선택"
        "shortcut_group" = "바로가기 옵션"
        "desktop_shortcut" = "바탕화면 바로가기 생성"
        "continue_button" = "계속"
        "cancel_button" = "취소"
        
        # 메인 폼
        "main_title" = "Open WebUI 설치 마법사"
        "status_ready" = "준비됨"
        "install_button" = "Open WebUI 설치 및 실행"
        "welcome_message" = "Open WebUI 설치 마법사에 오신 것을 환영합니다!"
        "help_message" = "이 프로그램은 Open WebUI를 쉽게 설치하고 실행할 수 있도록 도와줍니다."
        "start_message" = "시작하려면 'Open WebUI 설치 및 실행' 버튼을 클릭하세요."
        "admin_note" = "참고: 이 프로그램은 관리자 권한으로 실행됩니다."

        # 설치 과정 메시지
        "installation_start" = "Open WebUI 설치 및 실행 시작"
        "step1" = "1단계: Conda 설치 확인/설치"
        "step2" = "2단계: Open WebUI 설치 및 실행"
        "conda_check" = "Conda 설치 확인 중..."
        "conda_found" = "Conda가 이미 설치되어 있습니다: {0}"
        "miniconda_found" = "Miniconda가 이미 설치되어 있습니다: {0}"
        "conda_not_found" = "Conda가 설치되어 있지 않습니다. 설치를 시작합니다..."
        "miniconda_download" = "Miniconda 다운로드 중..."
        "miniconda_installing" = "Miniconda 설치 중..."
        "miniconda_success" = "Miniconda가 성공적으로 설치되었습니다."
        "miniconda_path" = "Miniconda 설치 경로: {0}"
        "miniconda_path_not_found" = "Miniconda 설치 경로를 찾을 수 없습니다."
        "miniconda_error" = "Miniconda 설치 중 오류가 발생했습니다. 종료 코드: {0}"
        "conda_download_error" = "Miniconda 다운로드 또는 설치 중 오류가 발생했습니다: {0}"
        "conda_check_error" = "Conda 설치 확인 중 오류가 발생했습니다: {0}"
        "conda_install_complete" = "Conda 설치 완료"
        "conda_install_failed" = "Conda 설치 실패"
        "conda_confirmed" = "Conda 설치가 확인되었습니다: {0}"

        # Open WebUI 설치 메시지
        "openwebui_install_start" = "Open WebUI 설치 및 실행 시작..."
        "conda_env_creating" = "1. Conda 환경 생성 중: open-webui"
        "env_exists" = "open-webui 환경이 이미 존재합니다."
        "env_creating" = "open-webui 환경을 생성합니다..."
        "env_activate" = "2. open-webui 환경 활성화 및 패키지 설치 중..."
        "package_installing" = "패키지 설치 중..."
        "activate_not_found" = "활성화 스크립트를 찾을 수 없습니다: {0}"
        "try_alternative" = "다른 경로를 시도합니다..."
        "alternative_found" = "대체 활성화 스크립트를 찾았습니다: {0}"
        "activate_not_found_abort" = "활성화 스크립트를 찾을 수 없습니다. 설치를 중단합니다."
        "file_setup" = "3. Open WebUI 파일 설정 중..."
        "file_setup_progress" = "파일 설정 중..."
        "folder_created" = "OpenWebui 폴더를 생성했습니다: {0}"
        "folder_exists" = "OpenWebui 폴더가 이미 존재합니다: {0}"
        "batch_created" = "start_openwebui.bat 파일을 생성했습니다."
        "vbs_created" = "start_openwebui.vbs 파일을 생성했습니다."
        "html_created" = "loading.html 파일을 생성했습니다."
        "shortcut_creating" = "4. 바로가기 파일 생성 중..."
        "creating_shortcuts" = "바로가기 생성 중..."
        "shortcut_created" = "Open-Webui 바로가기를 생성했습니다."
        "desktop_shortcut_created" = "바탕화면에 바로가기가 생성되었습니다."
        "alluser_shortcut_created" = "모든 사용자 Start Menu > Programs에 바로가기가 생성되었습니다."
        "user_shortcut_created" = "현재 사용자 Start Menu > Programs에 바로가기가 생성되었습니다."
        "desktop_shortcut_error" = "바탕화면 바로가기 생성 중 오류 발생: {0}"
        "alluser_shortcut_error" = "모든 사용자 Start Menu > Programs 복사 중 오류 발생: {0}"
        "user_shortcut_error" = "현재 사용자 Start Menu > Programs 복사 중 오류 발생: {0}"
        "launching" = "5. Open WebUI 실행 중..."
        "launching_progress" = "Open WebUI 실행 중..."
        "launch_success" = "Open WebUI가 시작되었습니다."
        "browser_message" = "웹 브라우저에서 http://localhost:8080/ 을 열어 접속하세요."
        "install_complete" = "설치 및 실행 완료!"
        "install_error" = "Open WebUI 설치 및 실행 중 오류가 발생했습니다: {0}"
        "install_process_error" = "설치 과정에서 오류가 발생했습니다: {0}"

        # 배치 파일 내용
        "batch_checking" = "Open WebUI 상태 확인 중..."
        "batch_server_running" = "서버가 이미 실행 중입니다. 브라우저를 열어요!"
        "batch_server_not_running" = "서버가 실행 중이 아닙니다. 시작합니다..."
        "batch_waiting" = "서버가 시작될 때까지 기다리는 중..."

        "auto_retry" = "설치 중 문제가 발생했습니다. 2초 후 자동으로 재시도합니다..."
        "retrying" = "자동 재시도 중..."
        "retry_error" = "재시도 중에도 오류가 발생했습니다: {0}"
    }
    "English" = @{
        "separator" = "==================================================="
        "openwebui_installing" = "Installing Open WebUI..."

        # Initial Setup Form
        "setup_title" = "Open WebUI Installation Wizard"
        "setup_description" = "Please select basic settings for installing Open WebUI."
        "language_group" = "Language Selection"
        "shortcut_group" = "Shortcut Options"
        "desktop_shortcut" = "Create desktop shortcut"
        "continue_button" = "Continue"
        "cancel_button" = "Cancel"

        # Main Form
        "main_title" = "Open WebUI Installation Wizard"
        "status_ready" = "Ready"
        "install_button" = "Install and Run Open WebUI"
        "welcome_message" = "Welcome to the Open WebUI Installation Wizard!"
        "help_message" = "This program helps you easily install and run Open WebUI."
        "start_message" = "Click the 'Install and Run Open WebUI' button to start."
        "admin_note" = "Note: This program runs with administrator privileges."

        # Installation Process Messages
        "installation_start" = "Starting Open WebUI Installation and Execution"
        "step1" = "Step 1: Checking/Installing Conda"
        "step2" = "Step 2: Installing and Running Open WebUI"
        "conda_check" = "Checking Conda installation..."
        "conda_found" = "Conda is already installed: {0}"
        "miniconda_found" = "Miniconda is already installed: {0}"
        "conda_not_found" = "Conda is not installed. Starting installation..."
        "miniconda_download" = "Downloading Miniconda..."
        "miniconda_installing" = "Installing Miniconda..."
        "miniconda_success" = "Miniconda has been successfully installed."
        "miniconda_path" = "Miniconda installation path: {0}"
        "miniconda_path_not_found" = "Miniconda installation path not found."
        "miniconda_error" = "An error occurred during Miniconda installation. Exit code: {0}"
        "conda_download_error" = "An error occurred during Miniconda download or installation: {0}"
        "conda_check_error" = "An error occurred while checking Conda installation: {0}"
        "conda_install_complete" = "Conda installation complete"
        "conda_install_failed" = "Conda installation failed"
        "conda_confirmed" = "Conda installation confirmed: {0}"

        # Open WebUI Installation Messages
        "openwebui_install_start" = "Starting Open WebUI installation and execution..."
        "conda_env_creating" = "1. Creating Conda environment: open-webui"
        "env_exists" = "open-webui environment already exists."
        "env_creating" = "Creating open-webui environment..."
        "env_activate" = "2. Activating open-webui environment and installing packages..."
        "package_installing" = "Installing packages..."
        "activate_not_found" = "Activation script not found: {0}"
        "try_alternative" = "Trying alternative path..."
        "alternative_found" = "Found alternative activation script: {0}"
        "activate_not_found_abort" = "Activation script not found. Aborting installation."
        "file_setup" = "3. Setting up Open WebUI files..."
        "file_setup_progress" = "Setting up files..."
        "folder_created" = "Created OpenWebui folder: {0}"
        "folder_exists" = "OpenWebui folder already exists: {0}"
        "batch_created" = "Created start_openwebui.bat file."
        "vbs_created" = "Created start_openwebui.vbs file."
        "html_created" = "Created loading.html file."
        "shortcut_creating" = "4. Creating shortcut files..."
        "creating_shortcuts" = "Creating shortcuts..."
        "shortcut_created" = "Created Open-Webui shortcut."
        "desktop_shortcut_created" = "Desktop shortcut created."
        "alluser_shortcut_created" = "Shortcut created in All Users Start Menu > Programs."
        "user_shortcut_created" = "Shortcut created in Current User Start Menu > Programs."
        "desktop_shortcut_error" = "Error creating desktop shortcut: {0}"
        "alluser_shortcut_error" = "Error copying to All Users Start Menu > Programs: {0}"
        "user_shortcut_error" = "Error copying to Current User Start Menu > Programs: {0}"
        "launching" = "5. Launching Open WebUI..."
        "launching_progress" = "Launching Open WebUI..."
        "launch_success" = "Open WebUI has been started."
        "browser_message" = "Access it by opening http://localhost:8080/ in your web browser."
        "install_complete" = "Installation and execution complete!"
        "install_error" = "An error occurred during Open WebUI installation and execution: {0}"
        "install_process_error" = "An error occurred during the installation process: {0}"

        # Batch File Content
        "batch_checking" = "Checking Open WebUI status..."
        "batch_server_running" = "Server is already running. Opening browser!"
        "batch_server_not_running" = "Server is not running. Starting..."
        "batch_waiting" = "Waiting for server to start..."

        "auto_retry" = "A problem occurred during installation. Automatically retrying in 2 seconds..."
        "retrying" = "Automatically retrying..."
        "retry_error" = "Error occurred during retry: {0}"
    }
}

# --- Get-Text 함수 -------------------------------------------------------------
function Get-Text {
    param([string]$key,[array]$params=@())
    $text = $global:translations[$global:selectedLanguage][$key]
    if ($params.Count -gt 0) {
        $text = [string]::Format($text, $params)
    }
    return $text
}

# --- 초기 설정 폼 (언어/바탕화면 바로가기) ---
function Show-InitialSetupForm {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    
    $initialForm = New-Object System.Windows.Forms.Form
    $initialForm.Text = Get-Text "setup_title"
    $initialForm.Size = New-Object System.Drawing.Size(520, 400)
    $initialForm.StartPosition = "CenterScreen"
    $initialForm.FormBorderStyle = 'FixedDialog'
    $initialForm.MaximizeBox = $false
    $initialForm.MinimizeBox = $false
    $initialForm.BackColor = [System.Drawing.Color]::White
    $initialForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    # 상단 헤더
    $headerPanel = New-Object System.Windows.Forms.Panel
    $headerPanel.Dock = 'Top'
    $headerPanel.Height = 80
    $headerPanel.BackColor = [System.Drawing.Color]::FromArgb(0, 99, 177)
    $initialForm.Controls.Add($headerPanel)

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = Get-Text "setup_title"
    $titleLabel.ForeColor = [System.Drawing.Color]::White
    $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
    $titleLabel.AutoSize = $false
    $titleLabel.Size = New-Object System.Drawing.Size(400, 40)
    $titleLabel.Location = New-Object System.Drawing.Point(20, 20)
    $titleLabel.TextAlign = 'MiddleLeft'
    $headerPanel.Controls.Add($titleLabel)

    # 설명 라벨
    $descriptionLabel = New-Object System.Windows.Forms.Label
    $descriptionLabel.Text = Get-Text "setup_description"
    $descriptionLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $descriptionLabel.ForeColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
    $descriptionLabel.Location = New-Object System.Drawing.Point(30, 90)
    $descriptionLabel.Size = New-Object System.Drawing.Size(460, 30)
    $initialForm.Controls.Add($descriptionLabel)

    # 구분선
    $separator = New-Object System.Windows.Forms.Panel
    $separator.BackColor = [System.Drawing.Color]::FromArgb(230, 230, 230)
    $separator.Location = New-Object System.Drawing.Point(30,120)
    $separator.Size = New-Object System.Drawing.Size(460, 1)
    $initialForm.Controls.Add($separator)

    # 언어 선택 그룹
    $languageGroup = New-Object System.Windows.Forms.GroupBox
    $languageGroup.Text = Get-Text "language_group"
    $languageGroup.Location = New-Object System.Drawing.Point(30,130)
    $languageGroup.Size = New-Object System.Drawing.Size(460, 90)
    $languageGroup.ForeColor = [System.Drawing.Color]::FromArgb(60,60,60)
    $languageGroup.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $initialForm.Controls.Add($languageGroup)

    $koreanRadio = New-Object System.Windows.Forms.RadioButton
    $koreanRadio.Text = "한국어"
    $koreanRadio.Location = New-Object System.Drawing.Point(30,35)
    $koreanRadio.Size = New-Object System.Drawing.Size(180,30)
    $koreanRadio.Checked = $true
    $koreanRadio.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $koreanRadio.Add_CheckedChanged({
        if($koreanRadio.Checked){ $global:selectedLanguage = "한국어" }
    })
    $languageGroup.Controls.Add($koreanRadio)

    $englishRadio = New-Object System.Windows.Forms.RadioButton
    $englishRadio.Text = "English"
    $englishRadio.Location = New-Object System.Drawing.Point(240,35)
    $englishRadio.Size = New-Object System.Drawing.Size(180,30)
    $englishRadio.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $englishRadio.Add_CheckedChanged({
        if($englishRadio.Checked){ $global:selectedLanguage = "English" }
    })
    $languageGroup.Controls.Add($englishRadio)

    # 바로가기 옵션 그룹
    $shortcutGroup = New-Object System.Windows.Forms.GroupBox
    $shortcutGroup.Text = Get-Text "shortcut_group"
    $shortcutGroup.Location = New-Object System.Drawing.Point(30,230)
    $shortcutGroup.Size = New-Object System.Drawing.Size(460,70)
    $shortcutGroup.ForeColor = [System.Drawing.Color]::FromArgb(60,60,60)
    $shortcutGroup.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $initialForm.Controls.Add($shortcutGroup)

    $desktopShortcutCheckbox = New-Object System.Windows.Forms.CheckBox
    $desktopShortcutCheckbox.Text = Get-Text "desktop_shortcut"
    $desktopShortcutCheckbox.Location = New-Object System.Drawing.Point(30,30)
    $desktopShortcutCheckbox.Size = New-Object System.Drawing.Size(400,25)
    $desktopShortcutCheckbox.Checked = $true
    $desktopShortcutCheckbox.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $desktopShortcutCheckbox.Add_CheckedChanged({
        $global:createDesktopShortcut = $desktopShortcutCheckbox.Checked
    })
    $shortcutGroup.Controls.Add($desktopShortcutCheckbox)

    # 하단 버튼 패널
    $buttonPanel = New-Object System.Windows.Forms.Panel
    $buttonPanel.Dock = 'Bottom'
    $buttonPanel.Height = 60
    $buttonPanel.BackColor = [System.Drawing.Color]::FromArgb(245,245,245)
    $initialForm.Controls.Add($buttonPanel)

    # 계속 버튼
    $continueButton = New-Object System.Windows.Forms.Button
    $continueButton.Text = Get-Text "continue_button"
    $continueButton.Location = New-Object System.Drawing.Point(370,15)
    $continueButton.Size = New-Object System.Drawing.Size(120,35)
    $continueButton.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $continueButton.BackColor = [System.Drawing.Color]::FromArgb(0,120,215)
    $continueButton.ForeColor = [System.Drawing.Color]::White
    $continueButton.FlatStyle = 'Flat'
    $continueButton.FlatAppearance.BorderSize = 0
    $continueButton.Cursor = 'Hand'
    $continueButton.Add_Click({
        $initialForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $initialForm.Close()
    })
    $buttonPanel.Controls.Add($continueButton)

    # 취소 버튼
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = Get-Text "cancel_button"
    $cancelButton.Location = New-Object System.Drawing.Point(240,15)
    $cancelButton.Size = New-Object System.Drawing.Size(120,35)
    $cancelButton.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $cancelButton.BackColor = [System.Drawing.Color]::FromArgb(240,240,240)
    $cancelButton.ForeColor = [System.Drawing.Color]::FromArgb(60,60,60)
    $cancelButton.FlatStyle = 'Flat'
    $cancelButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(210,210,210)
    $cancelButton.Cursor = 'Hand'
    $cancelButton.Add_Click({
        $initialForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $initialForm.Close()
    })
    $buttonPanel.Controls.Add($cancelButton)

    # 마우스 오버 효과 (PowerShell WinForms에서는 $this가 유효하지 않아 에러 가능성 있으나 원문 그대로 유지)
    $continueButton.Add_MouseEnter({ $this.BackColor = [System.Drawing.Color]::FromArgb(0,102,184) })
    $continueButton.Add_MouseLeave({ $this.BackColor = [System.Drawing.Color]::FromArgb(0,120,215) })
    $cancelButton.Add_MouseEnter({ $this.BackColor = [System.Drawing.Color]::FromArgb(230,230,230) })
    $cancelButton.Add_MouseLeave({ $this.BackColor = [System.Drawing.Color]::FromArgb(240,240,240) })

    # 초기 설정 값
    $global:selectedLanguage = "한국어"
    $global:createDesktopShortcut = $true

    return $initialForm.ShowDialog()
}

# 바로가기 관리 함수
function Manage-Shortcuts {
    param([string]$existingShortcutPath)

    if($global:createDesktopShortcut) {
        try {
            $desktopPath = [Environment]::GetFolderPath("Desktop")
            $desktopShortcut = Join-Path $desktopPath "Open WebUI.lnk"
            Copy-Item -Path $existingShortcutPath -Destination $desktopShortcut -Force
            Update-Log (Get-Text "desktop_shortcut_created") "green"
        } catch {
            Update-Log (Get-Text "desktop_shortcut_error" -params @($_)) "red"
        }
    }

    $allUserPrograms = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs"
    if(Test-Path $allUserPrograms) {
        try {
            $destination = Join-Path $allUserPrograms "Open WebUI.lnk"
            Copy-Item -Path $existingShortcutPath -Destination $destination -Force
            Update-Log (Get-Text "alluser_shortcut_created") "green"
        } catch {
            Update-Log (Get-Text "alluser_shortcut_error" -params @($_)) "red"
        }
    }

    $userPrograms = Join-Path $env:USERPROFILE "AppData\Roaming\Microsoft\Windows\Start Menu\Programs"
    if(Test-Path $userPrograms) {
        try {
            $destination = Join-Path $userPrograms "Open WebUI.lnk"
            Copy-Item -Path $existingShortcutPath -Destination $destination -Force
            Update-Log (Get-Text "user_shortcut_created") "green"
        } catch {
            Update-Log (Get-Text "user_shortcut_error" -params @($_)) "red"
        }
    }
}

# 인코딩 설정
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# 애플리케이션 정보
$appName = "Open WebUI"
$appVersion = "1.0.0"
$companyName = "Open WebUI Installer"

# 색상 테마
$primaryColor = [System.Drawing.Color]::FromArgb(30,80,162)
$secondaryColor = [System.Drawing.Color]::FromArgb(240,245,250)
$accentColor = [System.Drawing.Color]::FromArgb(0,120,215)
$textColor = [System.Drawing.Color]::FromArgb(50,50,50)
$lightTextColor = [System.Drawing.Color]::White

# 폰트 설정
$defaultFont = New-Object System.Drawing.Font("Segoe UI",9)
$headerFont  = New-Object System.Drawing.Font("Segoe UI",12,[System.Drawing.FontStyle]::Bold)
$boldFont    = New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)

# 전역 변수
$global:installedCondaPath = $null

# ---------------------------
# 메인 폼 만들기 (Build-MainForm)
# ---------------------------
function Build-MainForm {
    # 폼 생성
    $form = New-Object System.Windows.Forms.Form
    $form.Text = Get-Text "main_title"
    $form.Size = New-Object System.Drawing.Size(800,600)
    $form.StartPosition = "CenterScreen"
    $form.BackColor = $secondaryColor
    $form.Font = $defaultFont
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false

    # 헤더 패널
    $headerPanel = New-Object System.Windows.Forms.Panel
    $headerPanel.Dock = 'Top'
    $headerPanel.Height = 70
    $headerPanel.BackColor = $primaryColor
    $form.Controls.Add($headerPanel)

    # 타이틀 라벨
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = Get-Text "main_title"
    $titleLabel.ForeColor = $lightTextColor
    $titleLabel.Font = $headerFont
    $titleLabel.AutoSize = $true
    $titleLabel.Location = New-Object System.Drawing.Point(20,20)
    $headerPanel.Controls.Add($titleLabel)

    # 콘텐츠 패널
    $contentPanel = New-Object System.Windows.Forms.Panel
    $contentPanel.Dock = 'Fill'
    $contentPanel.BackColor = $secondaryColor
    $contentPanel.Padding = New-Object System.Windows.Forms.Padding(20)
    $form.Controls.Add($contentPanel)

    # 상태 패널
    $statusPanel = New-Object System.Windows.Forms.Panel
    $statusPanel.Dock = 'Bottom'
    $statusPanel.Height = 30
    $statusPanel.BackColor = [System.Drawing.Color]::FromArgb(230,230,230)
    $form.Controls.Add($statusPanel)

    # 상태 라벨
    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Text = Get-Text "status_ready"
    $statusLabel.AutoSize = $true
    $statusLabel.Location = New-Object System.Drawing.Point(10,7)
    $statusPanel.Controls.Add($statusLabel)

    # 버튼 패널
    $buttonPanel = New-Object System.Windows.Forms.Panel
    $buttonPanel.Dock = 'Bottom'
    $buttonPanel.Height = 60
    $buttonPanel.BackColor = [System.Drawing.Color]::White
    $form.Controls.Add($buttonPanel)

    # 진행 상태 표시줄
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Dock = 'Bottom'
    $progressBar.Height = 5
    $progressBar.Style = 'Continuous'
    $progressBar.MarqueeAnimationSpeed = 30
    $progressBar.Value = 0
    $form.Controls.Add($progressBar)

    # 로그 텍스트박스
    $txtLog = New-Object System.Windows.Forms.RichTextBox
    $txtLog.Dock = 'Fill'
    $txtLog.ReadOnly = $true
    $txtLog.BackColor = [System.Drawing.Color]::White
    $txtLog.Font = New-Object System.Drawing.Font("Consolas",9)
    $txtLog.BorderStyle = 'None'
    $contentPanel.Controls.Add($txtLog)

    # 전역 레퍼런스
    $global:txtLog = $txtLog
    $global:progressBar = $progressBar
    $global:statusLabel = $statusLabel

    # 단일 버튼: 설치 & 실행
    $installButton = New-Object System.Windows.Forms.Button
    $installButton.Text = Get-Text "install_button"
    $installButton.Width = 250
    $installButton.Height = 36
    $installButton.Location = New-Object System.Drawing.Point((($buttonPanel.Width - 250) / 2), 12)
    $installButton.BackColor = $accentColor
    $installButton.ForeColor = $lightTextColor
    $installButton.FlatStyle = 'Flat'
    $installButton.Font = $boldFont
    $installButton.Cursor = 'Hand'
    $buttonPanel.Controls.Add($installButton)

    # ---------------------------
    # 설치 버튼 이벤트
    # ---------------------------
    $installButton.Add_Click({
        # 버튼 비활성화 (중복 클릭 방지)
        $installButton.Enabled = $false
        
        Update-Log (Get-Text "separator") "blue"
        Update-Log (Get-Text "installation_start") "blue"
        Update-Log (Get-Text "separator") "blue"
        
        # 전역 변수로 재시도 상태 추적
        $global:retryAttempted = $false
        
        try {
            # 1단계: Conda 설치 확인
            Update-Log (Get-Text "step1") "blue"
            $condaPath = Check-CondaInstallation
            if ($condaPath) {
                Update-Log (Get-Text "conda_confirmed" -params @($condaPath)) "green"
                Update-Progress -value 50 -step (Get-Text "conda_install_complete")
                # 2단계: Open WebUI 설치
                Update-Log (Get-Text "step2") "blue"
                $success = Install-OpenWebUI -condaPath $condaPath
                
                # 설치 실패 시 자동 재시도
                if (-not $success -and -not $global:retryAttempted) {
                    $global:retryAttempted = $true
                    Update-Log "설치 중 문제가 발생했습니다. 2초 후 자동으로 재시도합니다..." "blue"
                    Start-Sleep -Seconds 2
                    
                    # 재시도 로그
                    Update-Log (Get-Text "separator") "blue"
                    Update-Log "자동 재시도 중..." "blue"
                    Update-Log (Get-Text "separator") "blue"
                    
                    # 다시 시작
                    Update-Log (Get-Text "step1") "blue"
                    $condaPath = Check-CondaInstallation
                    if ($condaPath) {
                        Update-Log (Get-Text "conda_confirmed" -params @($condaPath)) "green"
                        Update-Progress -value 50 -step (Get-Text "conda_install_complete")
                        Update-Log (Get-Text "step2") "blue"
                        Install-OpenWebUI -condaPath $condaPath
                    }
                }
            } else {
                Update-Log (Get-Text "conda_install_failed") "red"
                Update-Progress -value 0 -step (Get-Text "conda_install_failed")
                
                # Conda 설치 실패 시 자동 재시도
                if (-not $global:retryAttempted) {
                    $global:retryAttempted = $true
                    Update-Log "Conda 설치 중 문제가 발생했습니다. 2초 후 자동으로 재시도합니다..." "blue"
                    Start-Sleep -Seconds 2
                    
                    # 재시도 로그
                    Update-Log (Get-Text "separator") "blue"
                    Update-Log "자동 재시도 중..." "blue"
                    Update-Log (Get-Text "separator") "blue"
                    
                    # 다시 시작
                    Update-Log (Get-Text "step1") "blue"
                    $condaPath = Check-CondaInstallation
                    if ($condaPath) {
                        Update-Log (Get-Text "conda_confirmed" -params @($condaPath)) "green"
                        Update-Progress -value 50 -step (Get-Text "conda_install_complete")
                        Update-Log (Get-Text "step2") "blue"
                        Install-OpenWebUI -condaPath $condaPath
                    }
                }
            }
        } catch {
            Update-Log (Get-Text "install_process_error" -params @($_)) "red"
            
            # 예외 발생 시 자동 재시도
            if (-not $global:retryAttempted) {
                $global:retryAttempted = $true
                Update-Log "설치 중 예외가 발생했습니다. 2초 후 자동으로 재시도합니다..." "blue"
                Start-Sleep -Seconds 2
                
                # 재시도 로그
                Update-Log (Get-Text "separator") "blue"
                Update-Log "자동 재시도 중..." "blue"
                Update-Log (Get-Text "separator") "blue"
                
                # 다시 시작
                try {
                    Update-Log (Get-Text "step1") "blue"
                    $condaPath = Check-CondaInstallation
                    if ($condaPath) {
                        Update-Log (Get-Text "conda_confirmed" -params @($condaPath)) "green"
                        Update-Progress -value 50 -step (Get-Text "conda_install_complete")
                        Update-Log (Get-Text "step2") "blue"
                        Install-OpenWebUI -condaPath $condaPath
                    }
                } catch {
                    Update-Log "재시도 중에도 오류가 발생했습니다: $_" "red"
                }
            }
        } finally {
            # 작업 완료 후 버튼 다시 활성화
            $installButton.Enabled = $true
        }
    })
    # 폼 로드 이벤트
    $form.Add_Load({
        Update-Log (Get-Text "welcome_message") "blue"
        Update-Log (Get-Text "help_message") "blue"
        Update-Log (Get-Text "start_message") "blue"
        Update-Log ""
        Update-Log (Get-Text "admin_note") "#888888"
    })
    return $form
}

# ---------------------------
# Update-Log 등 필요한 함수
# ---------------------------
function Update-Log {
    param([string]$message,[string]$color="black")
    $txtLog.SelectionStart = $txtLog.TextLength
    $txtLog.SelectionLength = 0
    switch($color) {
        "red" { $txtLog.SelectionColor = [System.Drawing.Color]::Red }
        "green" { $txtLog.SelectionColor = [System.Drawing.Color]::Green }
        "blue" { $txtLog.SelectionColor = [System.Drawing.Color]::Blue }
        "#888888" { $txtLog.SelectionColor = [System.Drawing.Color]::FromArgb(136,136,136) }
        default { $txtLog.SelectionColor = [System.Drawing.Color]::Black }
    }
    $txtLog.AppendText("$message`r`n")
    $txtLog.ScrollToCaret()
}

function Update-Status {
    param([string]$status)
    $statusLabel.Text = $status
}

function Update-Progress {
    param([int]$value,[string]$step)
    $progressBar.Value = $value
    Update-Status $step
}

function Check-CondaInstallation {
    Update-Progress -value 10 -step (Get-Text "conda_check")
    Update-Status (Get-Text "conda_check")
    try {
        # 여러 가능한 Conda 경로 확인
        $possibleCondaPaths = @(
            (Join-Path $env:USERPROFILE "miniconda3\Scripts\conda.exe"),
            (Join-Path $env:USERPROFILE "AppData\Local\miniconda3\Scripts\conda.exe"),
            (Join-Path $env:USERPROFILE "Anaconda3\Scripts\conda.exe"),
            "C:\ProgramData\miniconda3\Scripts\conda.exe",
            "C:\ProgramData\Anaconda3\Scripts\conda.exe"
        )
        
        foreach ($condaPath in $possibleCondaPaths) {
            if (Test-Path $condaPath) {
                $basePath = Split-Path -Parent (Split-Path -Parent $condaPath)
                Update-Log (Get-Text "miniconda_found" -params @($basePath)) "green"
                return $basePath
            }
        }
        
        # 명령어로 확인
        $condaCommand = Get-Command conda -ErrorAction SilentlyContinue
        if ($condaCommand) {
            $condaExePath = $condaCommand.Source
            $basePath = Split-Path -Parent (Split-Path -Parent $condaExePath)
            Update-Log (Get-Text "conda_found" -params @($basePath)) "green"
            return $basePath
        }
        
        Update-Log (Get-Text "conda_not_found") "blue"
        Update-Progress -value 20 -step (Get-Text "miniconda_download")
        $tempDir = Join-Path $env:TEMP "OpenWebUIInstaller"
        if (-not (Test-Path $tempDir)) { 
            New-Item -ItemType Directory -Path $tempDir -Force | Out-Null 
        }
        
        Update-Log (Get-Text "miniconda_download")
        try {
            $wc = New-Object System.Net.WebClient
            $wc.DownloadFile("https://repo.anaconda.com/miniconda/Miniconda3-latest-Windows-x86_64.exe",
                             (Join-Path $tempDir "Miniconda3-latest-Windows-x86_64.exe"))
            
            Update-Log (Get-Text "miniconda_installing")
            Update-Progress -value 40 -step (Get-Text "miniconda_installing")
            
            $installArgs = "/S /InstallationType=JustMe /AddToPath=1 /RegisterPython=1"
            $process = Start-Process -FilePath (Join-Path $tempDir "Miniconda3-latest-Windows-x86_64.exe") -ArgumentList $installArgs -Wait -PassThru
            
            if ($process.ExitCode -eq 0) {
                Update-Log (Get-Text "miniconda_success") "green"
                
                # 환경 변수 갱신
                $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "User") + ";" + 
                            [System.Environment]::GetEnvironmentVariable("Path", "Machine")
                
                # 잠시 대기
                Start-Sleep -Seconds 3
                
                # 설치 후 여러 가능한 경로 확인
                $basePaths = @(
                    (Join-Path $env:USERPROFILE "miniconda3"),
                    (Join-Path $env:USERPROFILE "AppData\Local\miniconda3"),
                    "C:\ProgramData\miniconda3"
                )
                
                foreach ($path in $basePaths) {
                    if (Test-Path $path) {
                        Update-Log (Get-Text "miniconda_path" -params @($path)) "green"
                        return $path
                    }
                }
                
                # 기본 경로 반환 (일반적으로 가장 많이 사용되는 경로)
                $defaultPath = Join-Path $env:USERPROFILE "miniconda3"
                Update-Log "기본 Miniconda 경로 사용: $defaultPath" "blue"
                return $defaultPath
            } else {
                Update-Log (Get-Text "miniconda_error" -params @($process.ExitCode)) "red"
                return $null
            }
        } catch {
            Update-Log (Get-Text "conda_download_error" -params @($_)) "red"
            return $null
        }
    } catch {
        Update-Log (Get-Text "conda_check_error" -params @($_)) "red"
        return $null
    }
}

function Install-OpenWebUI {
    param([string]$condaPath)
    Update-Progress -value 60 -step (Get-Text "openwebui_installing")
    Update-Log (Get-Text "openwebui_install_start") "blue"
    try {
        # 1. Conda 환경 생성
        Update-Log (Get-Text "conda_env_creating") "blue"
        $envExists = $false
        $envList = & conda env list 2>&1
        foreach($line in $envList){
            if($line -match "open-webui"){
                Update-Log (Get-Text "env_exists") "green"
                $envExists = $true
                break
            }
        }
        if(-not $envExists){
            Update-Log (Get-Text "env_creating") "blue"
            $output = & conda create -n open-webui python=3.11 -y 2>&1
            foreach($line in $output){
                Update-Log $line
            }
        }
        # 2. 환경 활성화 & 패키지 설치
        Update-Log (Get-Text "env_activate") "blue"
        Update-Progress -value 70 -step (Get-Text "package_installing")
        
        # 여러 가능한 activate.bat 경로 시도
        $activateScriptPaths = @(
            (Join-Path $condaPath "Scripts\activate.bat"),
            (Join-Path (Split-Path -Parent $condaPath) "Scripts\activate.bat"),
            (Join-Path $env:USERPROFILE "AppData\Local\miniconda3\Scripts\activate.bat"),
            (Join-Path $env:USERPROFILE "miniconda3\Scripts\activate.bat"),
            "C:\ProgramData\miniconda3\Scripts\activate.bat"
        )
        
        $activateScript = $null
        foreach ($path in $activateScriptPaths) {
            if (Test-Path $path) {
                Update-Log "활성화 스크립트를 찾았습니다: $path" "green"
                $activateScript = $path
                break
            }
        }
        
        if (-not $activateScript) {
            Update-Log (Get-Text "activate_not_found_abort") "red"
            return $false
        }
        
        $installCommands = @"
call "$activateScript" open-webui
pip install open-webui
"@
        $tempBatchFile = Join-Path $env:TEMP "open_webui_install_temp.bat"
        Set-Content -Path $tempBatchFile -Value $installCommands -Encoding ASCII
        $output = & cmd.exe /c $tempBatchFile 2>&1
        foreach($line in $output){
            Update-Log $line
        }
        # 3. C:\OpenWebui 폴더 생성 및 파일 설정
        Update-Log (Get-Text "file_setup") "blue"
        Update-Progress -value 80 -step (Get-Text "file_setup_progress")
        $openWebUIFolder = "C:\OpenWebui"
        if(-not (Test-Path $openWebUIFolder)){
            New-Item -Path $openWebUIFolder -ItemType Directory -Force | Out-Null
            Update-Log (Get-Text "folder_created" -params @($openWebUIFolder)) "green"
        } else {
            Update-Log (Get-Text "folder_exists" -params @($openWebUIFolder)) "green"
        }
        # 배치 파일
        $batchContent = @"
@echo off
echo $(Get-Text 'batch_checking')
REM 서버가 이미 실행 중인지 확인
curl --silent --head http://localhost:8080/ >nul 2>&1
if %errorlevel% equ 0 (
    echo $(Get-Text "batch_server_running")
    start "" http://localhost:8080/
    exit
)
REM 서버가 실행 중이 아니면 로딩 화면 실행
echo $(Get-Text "batch_server_not_running")
start "" "loading.html"
REM Conda 환경 활성화 & 서버 시작
call "$activateScript" open-webui
start /B open-webui serve
echo $(Get-Text "batch_waiting")
exit
"@
        Set-Content -Path "$openWebUIFolder\start_openwebui.bat" -Value $batchContent -Force
        Update-Log (Get-Text "batch_created") "green"
        # vbs
        $vbsContent = @"
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
' 현재 스크립트가 위치한 폴더 경로
strPath = objFSO.GetParentFolderName(WScript.ScriptFullName)
objShell.CurrentDirectory = strPath
objShell.Run "start_openwebui.bat", 0, False
"@
        Set-Content -Path "$openWebUIFolder\start_openwebui.vbs" -Value $vbsContent -Force
        Update-Log (Get-Text "vbs_created") "green"
        # loading.html
        $htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <title>Open WebUI 로딩 중...</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            height: 100vh;
            margin: 0;
            background-color: #f5f5f5;
        }
        .loader {
            border: 16px solid #f3f3f3;
            border-radius: 50%;
            border-top: 16px solid #3498db;
            width: 120px;
            height: 120px;
            animation: spin 2s linear infinite;
            margin-bottom: 30px;
        }
        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
        h1 {
            color: #333;
        }
        #status {
            margin-top: 20px;
            color: #666;
        }
        #progress {
            margin-top: 10px;
            width: 300px;
            text-align: center;
        }
    </style>
    <script>
        let dots = 0;
        let checkCount = 0;
        function updateDots() {
            dots = (dots + 1) % 4;
            let dotText = '.'.repeat(dots);
            document.getElementById('dots').textContent = dotText;
        }
        function checkServer() {
            checkCount++;
            document.getElementById('progress').textContent = '서버 연결 시도 중...';
            fetch('http://localhost:8080/', { method: 'HEAD' })
                .then(response => {
                    if (response.ok) {
                        document.getElementById('status').textContent = '서버가 준비되었습니다! 잠시 후 자동으로 이동합니다.';
                        setTimeout(() => {
                            window.location.href = 'http://localhost:8080/';
                        }, 1500);
                    } else {
                        setTimeout(checkServer, 2000);
                    }
                })
                .catch(error => {
                    setTimeout(checkServer, 2000);
                });
        }
        window.onload = function() {
            setInterval(updateDots, 500);
            checkServer();
        };
    </script>
</head>
<body>
    <div class="loader"></div>
    <h1>Open WebUI 시작 중<span id="dots"></span></h1>
    <div id="status">새 서버 인스턴스를 시작하는 중입니다. 잠시만 기다려주세요.</div>
    <div id="progress">서버 연결 시도 중...</div>
</body>
</html>
"@
        Set-Content -Path "$openWebUIFolder\loading.html" -Value $htmlContent -Force
        Update-Log (Get-Text "html_created") "green"
        # 4. 바로가기 파일 생성
        Update-Log (Get-Text "shortcut_creating") "blue"
        Update-Progress -value 90 -step (Get-Text "creating_shortcuts")
        $WshShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut("$openWebUIFolder\Open-Webui.lnk")
        $Shortcut.TargetPath = "wscript.exe"
        $Shortcut.Arguments = "$openWebUIFolder\start_openwebui.vbs"
        $Shortcut.WorkingDirectory = "C:\OpenWebui"
        $Shortcut.Description = "Open WebUI 서버 실행"
        $Shortcut.Save()
        Update-Log (Get-Text "shortcut_created") "green"
        # 바탕화면/시작메뉴 복사
        $existingShortcutPath = "$openWebUIFolder\Open-Webui.lnk"
        Manage-Shortcuts -existingShortcutPath $existingShortcutPath
        # 5. 바로가기 실행
        Update-Log (Get-Text "launching") "blue"
        Update-Progress -value 95 -step (Get-Text "launching_progress")
        Start-Process "$openWebUIFolder\Open-Webui.lnk"
        Update-Log (Get-Text "launch_success") "green"
        Update-Log (Get-Text "browser_message") "green"
        Update-Progress -value 100 -step (Get-Text "install_complete")
        
        return $true  # 성공적으로 완료
    } catch {
        Update-Log (Get-Text "install_error" -params @($_)) "red"
        return $false  # 오류 발생 시 false 반환
    }
}

# ---------------------------
# 실제 실행 흐름
# ---------------------------
# 1) 초기 설정 폼 (언어/바탕화면) 먼저 열기
$initialResult = Show-InitialSetupForm

# 2) '취소'면 종료
if($initialResult -ne [System.Windows.Forms.DialogResult]::OK){
    [System.Windows.Forms.MessageBox]::Show(
        (Get-Text "cancel_button"),
        (Get-Text "setup_title"),
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
    exit
}

# 3) 초기 설정 끝났으면 메인 폼 생성 & 실행
$mainForm = Build-MainForm
[void]$mainForm.ShowDialog()
