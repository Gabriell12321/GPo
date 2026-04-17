# Auto-elevar como administrador se necessario
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Start-Process powershell -Verb RunAs -WindowStyle Hidden -ArgumentList "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$PSCommandPath`""
    exit
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ══════════════════════════════════════════
#  STATE
# ══════════════════════════════════════════
$script:Credential = $null
$script:Domain = ""
$script:UserName = ""

# Auto-detectar dominio
try {
    $script:Domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain().Name
} catch {
    $script:Domain = $env:USERDNSDOMAIN
    if ([string]::IsNullOrEmpty($script:Domain)) { $script:Domain = "" }
}
# Auto-detectar usuario
$script:UserName = $env:USERNAME
if ([string]::IsNullOrEmpty($script:UserName)) { $script:UserName = "" }

# ══════════════════════════════════════════
#  TEMA
# ══════════════════════════════════════════
$BgDark  = [System.Drawing.Color]::FromArgb(30, 30, 46)
$BgPanel = [System.Drawing.Color]::FromArgb(49, 50, 68)
$BgField = [System.Drawing.Color]::FromArgb(69, 71, 90)
$FgText  = [System.Drawing.Color]::FromArgb(205, 214, 244)
$Accent  = [System.Drawing.Color]::FromArgb(137, 180, 250)
$Green   = [System.Drawing.Color]::FromArgb(166, 227, 161)
$Red     = [System.Drawing.Color]::FromArgb(243, 139, 168)
$Yellow  = [System.Drawing.Color]::FromArgb(249, 226, 175)
$Mauve   = [System.Drawing.Color]::FromArgb(203, 166, 247)
$Overlay = [System.Drawing.Color]::FromArgb(108, 112, 134)

# ══════════════════════════════════════════
#  CONFIGURACOES PERSISTENTES
# ══════════════════════════════════════════
$script:SettingsPath = Join-Path $PSScriptRoot "gpo_settings.json"
function Load-Settings {
    if (Test-Path $script:SettingsPath) {
        try { return Get-Content $script:SettingsPath -Raw | ConvertFrom-Json } catch {}
    }
    return @{}
}
function Save-Settings {
    param($settings)
    $settings | ConvertTo-Json -Depth 5 | Set-Content $script:SettingsPath -Encoding UTF8 -Force
}
$script:Settings = Load-Settings

# ══════════════════════════════════════════
#  MESSAGEBOXES DARK (substitui MessageBox padrao)
# ══════════════════════════════════════════
function Show-DarkMsg {
    param(
        [string]$Message,
        [string]$Title = "Aviso",
        [string]$Buttons = "OK",
        [string]$Icon = "Information"
    )
    $d = New-Object System.Windows.Forms.Form
    $d.Text = $Title
    $d.Size = New-Object System.Drawing.Size(480, 220)
    $d.StartPosition = "CenterParent"
    $d.BackColor = $BgPanel; $d.ForeColor = $FgText
    $d.FormBorderStyle = "FixedDialog"
    $d.MaximizeBox = $false; $d.MinimizeBox = $false
    $d.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $d.ShowInTaskbar = $false

    # Icone
    $iconLbl = New-Object System.Windows.Forms.Label
    $iconLbl.Font = New-Object System.Drawing.Font("Segoe UI", 28)
    $iconLbl.AutoSize = $true
    $iconLbl.Location = New-Object System.Drawing.Point(20, 20)
    switch ($Icon) {
        "Information" { $iconLbl.Text = [char]0x2139; $iconLbl.ForeColor = $Accent }
        "Warning"     { $iconLbl.Text = [char]0x26A0; $iconLbl.ForeColor = $Yellow }
        "Error"       { $iconLbl.Text = [char]0x2716; $iconLbl.ForeColor = $Red }
        "Question"    { $iconLbl.Text = "?"; $iconLbl.ForeColor = $Mauve }
        default       { $iconLbl.Text = [char]0x2139; $iconLbl.ForeColor = $Accent }
    }
    $d.Controls.Add($iconLbl)

    $msgLbl = New-Object System.Windows.Forms.Label
    $msgLbl.Text = $Message
    $msgLbl.AutoSize = $false
    $msgLbl.Size = New-Object System.Drawing.Size(370, 100)
    $msgLbl.Location = New-Object System.Drawing.Point(75, 20)
    $msgLbl.ForeColor = $FgText
    $d.Controls.Add($msgLbl)

    $btnPanel = New-Object System.Windows.Forms.Panel
    $btnPanel.Dock = "Bottom"; $btnPanel.Height = 55
    $btnPanel.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 58)
    $d.Controls.Add($btnPanel)

    if ($Buttons -eq "YesNo") {
        $bYes = New-Object System.Windows.Forms.Button
        $bYes.Text = "Sim"; $bYes.Size = New-Object System.Drawing.Size(100, 35)
        $bYes.Location = New-Object System.Drawing.Point(140, 10)
        $bYes.BackColor = $Green; $bYes.ForeColor = $BgDark; $bYes.FlatStyle = "Flat"
        $bYes.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        $bYes.Cursor = [System.Windows.Forms.Cursors]::Hand
        $bYes.DialogResult = "Yes"
        $btnPanel.Controls.Add($bYes)
        $d.AcceptButton = $bYes

        $bNo = New-Object System.Windows.Forms.Button
        $bNo.Text = "Nao"; $bNo.Size = New-Object System.Drawing.Size(100, 35)
        $bNo.Location = New-Object System.Drawing.Point(250, 10)
        $bNo.BackColor = $BgField; $bNo.ForeColor = $FgText; $bNo.FlatStyle = "Flat"
        $bNo.Cursor = [System.Windows.Forms.Cursors]::Hand
        $bNo.DialogResult = "No"
        $btnPanel.Controls.Add($bNo)
        $d.CancelButton = $bNo
    } else {
        $bOk = New-Object System.Windows.Forms.Button
        $bOk.Text = "OK"; $bOk.Size = New-Object System.Drawing.Size(120, 35)
        $bOk.Location = New-Object System.Drawing.Point(175, 10)
        if ($Icon -eq "Error") { $bOk.BackColor = $Red; $bOk.ForeColor = [System.Drawing.Color]::White }
        elseif ($Icon -eq "Warning") { $bOk.BackColor = $Yellow; $bOk.ForeColor = $BgDark }
        else { $bOk.BackColor = $Accent; $bOk.ForeColor = $BgDark }
        $bOk.FlatStyle = "Flat"
        $bOk.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        $bOk.Cursor = [System.Windows.Forms.Cursors]::Hand
        $bOk.DialogResult = "OK"
        $btnPanel.Controls.Add($bOk)
        $d.AcceptButton = $bOk
    }

    return $d.ShowDialog()
}

# ══════════════════════════════════════════
#  FORM PRINCIPAL
# ══════════════════════════════════════════
$form = New-Object System.Windows.Forms.Form
$form.Text = "GPO - Sistema de TI"
$form.Size = New-Object System.Drawing.Size(1250, 720)
$form.StartPosition = "CenterScreen"
$form.BackColor = $BgDark
$form.ForeColor = $FgText
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$form.MinimumSize = New-Object System.Drawing.Size(850, 550)
$form.KeyPreview = $true

# Tooltip provider global
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.BackColor = $BgPanel
$toolTip.ForeColor = $FgText
$toolTip.InitialDelay = 400
$toolTip.ReshowDelay = 200

# ══════════════════════════════════════════
#  LOGIN PANEL
# ══════════════════════════════════════════
$loginPanel = New-Object System.Windows.Forms.Panel
$loginPanel.Size = New-Object System.Drawing.Size(400, 430)
$loginPanel.BackColor = $BgPanel

function Center-LoginPanel {
    $loginPanel.Location = New-Object System.Drawing.Point(
        [int](($form.ClientSize.Width - $loginPanel.Width) / 2),
        [int](($form.ClientSize.Height - $loginPanel.Height) / 2)
    )
}

# Titulo
$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "GPO - Sistema de TI"
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
$lblTitle.ForeColor = $Accent
$lblTitle.AutoSize = $true
$lblTitle.Location = New-Object System.Drawing.Point(65, 20)
$loginPanel.Controls.Add($lblTitle)

# Campo Dominio
$lblDom = New-Object System.Windows.Forms.Label
$lblDom.Text = "Dominio:"
$lblDom.ForeColor = $FgText
$lblDom.AutoSize = $true
$lblDom.Location = New-Object System.Drawing.Point(40, 75)
$loginPanel.Controls.Add($lblDom)

$txtDomain = New-Object System.Windows.Forms.TextBox
$txtDomain.Text = $script:Domain
$txtDomain.Location = New-Object System.Drawing.Point(40, 98)
$txtDomain.Size = New-Object System.Drawing.Size(320, 28)
$txtDomain.BackColor = $BgField
$txtDomain.ForeColor = $FgText
$txtDomain.Font = New-Object System.Drawing.Font("Segoe UI", 11)
$txtDomain.BorderStyle = "FixedSingle"
$loginPanel.Controls.Add($txtDomain)

$lblAutoDetect = New-Object System.Windows.Forms.Label
$lblAutoDetect.ForeColor = if ($script:Domain) { $Green } else { $Yellow }
$lblAutoDetect.Text = if ($script:Domain) { "(detectado automaticamente)" } else { "(nao detectado - digite manualmente)" }
$lblAutoDetect.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$lblAutoDetect.AutoSize = $true
$lblAutoDetect.Location = New-Object System.Drawing.Point(40, 126)
$loginPanel.Controls.Add($lblAutoDetect)

# Campo Usuario
$lblUser = New-Object System.Windows.Forms.Label
$lblUser.Text = "Usuario:"
$lblUser.ForeColor = $FgText
$lblUser.AutoSize = $true
$lblUser.Location = New-Object System.Drawing.Point(40, 155)
$loginPanel.Controls.Add($lblUser)

$txtUser = New-Object System.Windows.Forms.TextBox
$txtUser.Text = $script:UserName
$txtUser.Location = New-Object System.Drawing.Point(40, 178)
$txtUser.Size = New-Object System.Drawing.Size(320, 28)
$txtUser.BackColor = $BgField
$txtUser.ForeColor = $FgText
$txtUser.Font = New-Object System.Drawing.Font("Segoe UI", 11)
$txtUser.BorderStyle = "FixedSingle"
$loginPanel.Controls.Add($txtUser)

# Campo Senha
$lblPass = New-Object System.Windows.Forms.Label
$lblPass.Text = "Senha:"
$lblPass.ForeColor = $FgText
$lblPass.AutoSize = $true
$lblPass.Location = New-Object System.Drawing.Point(40, 225)
$loginPanel.Controls.Add($lblPass)

$txtPass = New-Object System.Windows.Forms.TextBox
$txtPass.Location = New-Object System.Drawing.Point(40, 248)
$txtPass.Size = New-Object System.Drawing.Size(275, 28)
$txtPass.BackColor = $BgField
$txtPass.ForeColor = $FgText
$txtPass.Font = New-Object System.Drawing.Font("Segoe UI", 11)
$txtPass.BorderStyle = "FixedSingle"
$txtPass.UseSystemPasswordChar = $true
$loginPanel.Controls.Add($txtPass)

# Botao mostrar/ocultar senha
$btnShowPass = New-Object System.Windows.Forms.Button
$btnShowPass.Text = [char]0x25CF
$btnShowPass.Location = New-Object System.Drawing.Point(320, 248)
$btnShowPass.Size = New-Object System.Drawing.Size(40, 28)
$btnShowPass.BackColor = $BgField; $btnShowPass.ForeColor = $FgText
$btnShowPass.FlatStyle = "Flat"; $btnShowPass.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnShowPass.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$btnShowPass.TabStop = $false
$toolTip.SetToolTip($btnShowPass, "Mostrar/ocultar senha")
$btnShowPass.Add_Click({
    $txtPass.UseSystemPasswordChar = -not $txtPass.UseSystemPasswordChar
    $btnShowPass.Text = if ($txtPass.UseSystemPasswordChar) { [char]0x25CF } else { "abc" }
})
$loginPanel.Controls.Add($btnShowPass)

# Botao Entrar
$btnLogin = New-Object System.Windows.Forms.Button
$btnLogin.Text = "Entrar"
$btnLogin.Location = New-Object System.Drawing.Point(130, 300)
$btnLogin.Size = New-Object System.Drawing.Size(140, 40)
$btnLogin.BackColor = $Accent
$btnLogin.ForeColor = $BgDark
$btnLogin.FlatStyle = "Flat"
$btnLogin.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$btnLogin.Cursor = [System.Windows.Forms.Cursors]::Hand
$loginPanel.Controls.Add($btnLogin)

# Mensagem
$lblMsg = New-Object System.Windows.Forms.Label
$lblMsg.Text = ""
$lblMsg.AutoSize = $true
$lblMsg.MaximumSize = New-Object System.Drawing.Size(320, 50)
$lblMsg.Location = New-Object System.Drawing.Point(40, 355)
$loginPanel.Controls.Add($lblMsg)

$form.Controls.Add($loginPanel)

# ══════════════════════════════════════════
#  DASHBOARD PANEL
# ══════════════════════════════════════════
$dashPanel = New-Object System.Windows.Forms.Panel
$dashPanel.Dock = "Fill"
$dashPanel.BackColor = $BgDark
$dashPanel.Visible = $false

# Header
$header = New-Object System.Windows.Forms.Panel
$header.Dock = "Top"
$header.Height = 55
$header.BackColor = $BgPanel

$lblHeader = New-Object System.Windows.Forms.Label
$lblHeader.Text = "GPO Manager"
$lblHeader.Font = New-Object System.Drawing.Font("Segoe UI", 15, [System.Drawing.FontStyle]::Bold)
$lblHeader.ForeColor = $Accent
$lblHeader.AutoSize = $true
$lblHeader.Location = New-Object System.Drawing.Point(15, 14)
$header.Controls.Add($lblHeader)

$lblVersion = New-Object System.Windows.Forms.Label
$lblVersion.Text = "v2.0"
$lblVersion.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$lblVersion.ForeColor = $Overlay
$lblVersion.AutoSize = $true
$lblVersion.Location = New-Object System.Drawing.Point(175, 24)
$header.Controls.Add($lblVersion)

$lblInfo = New-Object System.Windows.Forms.Label
$lblInfo.ForeColor = $FgText
$lblInfo.AutoSize = $true
$lblInfo.Location = New-Object System.Drawing.Point(250, 18)
$header.Controls.Add($lblInfo)

$btnLogout = New-Object System.Windows.Forms.Button
$btnLogout.Text = "Sair"
$btnLogout.Size = New-Object System.Drawing.Size(70, 32)
$btnLogout.BackColor = $Red
$btnLogout.ForeColor = [System.Drawing.Color]::White
$btnLogout.FlatStyle = "Flat"
$btnLogout.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnLogout.Anchor = "Top, Right"
$header.Controls.Add($btnLogout)

$dashPanel.Controls.Add($header)

# Toolbar GPO
$toolbar = New-Object System.Windows.Forms.Panel
$toolbar.Dock = "Top"
$toolbar.Height = 50
$toolbar.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 58)

$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Text = "Atualizar"
$btnRefresh.Location = New-Object System.Drawing.Point(10, 10)
$btnRefresh.Size = New-Object System.Drawing.Size(110, 30)
$btnRefresh.BackColor = $Accent
$btnRefresh.ForeColor = $BgDark
$btnRefresh.FlatStyle = "Flat"
$btnRefresh.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnRefresh.Cursor = [System.Windows.Forms.Cursors]::Hand
$toolTip.SetToolTip($btnRefresh, "Atualizar lista de GPOs (F5)")
$toolbar.Controls.Add($btnRefresh)

$btnNewGPO = New-Object System.Windows.Forms.Button
$btnNewGPO.Text = "Nova GPO"
$btnNewGPO.Location = New-Object System.Drawing.Point(130, 10)
$btnNewGPO.Size = New-Object System.Drawing.Size(110, 30)
$btnNewGPO.BackColor = $Green
$btnNewGPO.ForeColor = $BgDark
$btnNewGPO.FlatStyle = "Flat"
$btnNewGPO.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnNewGPO.Cursor = [System.Windows.Forms.Cursors]::Hand
$toolTip.SetToolTip($btnNewGPO, "Criar nova GPO vazia (Ctrl+N)")
$toolbar.Controls.Add($btnNewGPO)

$btnEditGPO = New-Object System.Windows.Forms.Button
$btnEditGPO.Text = "Editar GPO"
$btnEditGPO.Location = New-Object System.Drawing.Point(250, 10)
$btnEditGPO.Size = New-Object System.Drawing.Size(110, 30)
$btnEditGPO.BackColor = $Yellow
$btnEditGPO.ForeColor = $BgDark
$btnEditGPO.FlatStyle = "Flat"
$btnEditGPO.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnEditGPO.Cursor = [System.Windows.Forms.Cursors]::Hand
$toolTip.SetToolTip($btnEditGPO, "Editar GPO selecionada (Enter)")
$toolbar.Controls.Add($btnEditGPO)

$btnDeleteGPO = New-Object System.Windows.Forms.Button
$btnDeleteGPO.Text = "Excluir GPO"
$btnDeleteGPO.Location = New-Object System.Drawing.Point(370, 10)
$btnDeleteGPO.Size = New-Object System.Drawing.Size(110, 30)
$btnDeleteGPO.BackColor = $Red
$btnDeleteGPO.ForeColor = [System.Drawing.Color]::White
$btnDeleteGPO.FlatStyle = "Flat"
$btnDeleteGPO.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnDeleteGPO.Cursor = [System.Windows.Forms.Cursors]::Hand
$toolTip.SetToolTip($btnDeleteGPO, "Excluir GPO selecionada (Del)")
$toolbar.Controls.Add($btnDeleteGPO)

$btnWizard = New-Object System.Windows.Forms.Button
$btnWizard.Text = "⚡ GPO Rapida"
$btnWizard.Location = New-Object System.Drawing.Point(490, 10)
$btnWizard.Size = New-Object System.Drawing.Size(130, 30)
$btnWizard.BackColor = [System.Drawing.Color]::FromArgb(203,166,247)
$btnWizard.ForeColor = $BgDark
$btnWizard.FlatStyle = "Flat"
$btnWizard.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnWizard.Cursor = [System.Windows.Forms.Cursors]::Hand
$toolTip.SetToolTip($btnWizard, "Assistente de criacao rapida (Ctrl+W)")
$toolbar.Controls.Add($btnWizard)

$btnCloneGPO = New-Object System.Windows.Forms.Button
$btnCloneGPO.Text = "Duplicar"
$btnCloneGPO.Location = New-Object System.Drawing.Point(630, 10)
$btnCloneGPO.Size = New-Object System.Drawing.Size(90, 30)
$btnCloneGPO.BackColor = $BgField
$btnCloneGPO.ForeColor = $FgText
$btnCloneGPO.FlatStyle = "Flat"
$btnCloneGPO.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnCloneGPO.Cursor = [System.Windows.Forms.Cursors]::Hand
$toolTip.SetToolTip($btnCloneGPO, "Duplicar GPO selecionada (Ctrl+D)")
$toolbar.Controls.Add($btnCloneGPO)

$btnExportGPO = New-Object System.Windows.Forms.Button
$btnExportGPO.Text = "Exportar"
$btnExportGPO.Location = New-Object System.Drawing.Point(730, 10)
$btnExportGPO.Size = New-Object System.Drawing.Size(90, 30)
$btnExportGPO.BackColor = $BgField
$btnExportGPO.ForeColor = $FgText
$btnExportGPO.FlatStyle = "Flat"
$btnExportGPO.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnExportGPO.Cursor = [System.Windows.Forms.Cursors]::Hand
$toolTip.SetToolTip($btnExportGPO, "Exportar GPO para arquivo (Ctrl+E)")
$toolbar.Controls.Add($btnExportGPO)

$btnImportGPO = New-Object System.Windows.Forms.Button
$btnImportGPO.Text = "Importar"
$btnImportGPO.Location = New-Object System.Drawing.Point(830, 10)
$btnImportGPO.Size = New-Object System.Drawing.Size(90, 30)
$btnImportGPO.BackColor = $BgField
$btnImportGPO.ForeColor = $FgText
$btnImportGPO.FlatStyle = "Flat"
$btnImportGPO.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnImportGPO.Cursor = [System.Windows.Forms.Cursors]::Hand
$toolTip.SetToolTip($btnImportGPO, "Importar GPO de arquivo (Ctrl+I)")
$toolbar.Controls.Add($btnImportGPO)

$txtSearch = New-Object System.Windows.Forms.TextBox
$txtSearch.Location = New-Object System.Drawing.Point(940, 12)
$txtSearch.Size = New-Object System.Drawing.Size(200, 26)
$txtSearch.BackColor = $BgField
$txtSearch.ForeColor = $FgText
$txtSearch.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$txtSearch.BorderStyle = "FixedSingle"
$txtSearch.Text = ""
$txtSearch.Anchor = "Top, Right"
$toolbar.Controls.Add($txtSearch)

$lblSearch = New-Object System.Windows.Forms.Label
$lblSearch.Text = "Buscar:"
$lblSearch.ForeColor = $FgText
$lblSearch.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblSearch.AutoSize = $true
$lblSearch.Location = New-Object System.Drawing.Point(940, -2)
$lblSearch.Font = New-Object System.Drawing.Font("Segoe UI", 7)
$lblSearch.Anchor = "Top, Right"
$toolbar.Controls.Add($lblSearch)

$lblGPOCount = New-Object System.Windows.Forms.Label
$lblGPOCount.ForeColor = $Yellow
$lblGPOCount.AutoSize = $true
$lblGPOCount.Location = New-Object System.Drawing.Point(1150, 16)
$lblGPOCount.Anchor = "Top, Right"
$toolbar.Controls.Add($lblGPOCount)

$dashPanel.Controls.Add($toolbar)

# ── TabControl principal: GPOs | Cadastro de Apps ──
$mainTabs = New-Object System.Windows.Forms.TabControl
$mainTabs.Dock = "Fill"
$mainTabs.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$mainTabs.Padding = New-Object System.Drawing.Point(15, 6)

$tabGPOs = New-Object System.Windows.Forms.TabPage
$tabGPOs.Text = "  GPOs  "
$tabGPOs.BackColor = $BgDark

$tabApps = New-Object System.Windows.Forms.TabPage
$tabApps.Text = "  Cadastro de Apps  "
$tabApps.BackColor = $BgDark

# DataGridView GPOs
$dgv = New-Object System.Windows.Forms.DataGridView
$dgv.Dock = "Fill"
$dgv.BackgroundColor = $BgDark
$dgv.ForeColor = $FgText
$dgv.GridColor = $BgField
$dgv.BorderStyle = "None"
$dgv.CellBorderStyle = "SingleHorizontal"
$dgv.ColumnHeadersDefaultCellStyle.BackColor = $BgPanel
$dgv.ColumnHeadersDefaultCellStyle.ForeColor = $Accent
$dgv.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$dgv.ColumnHeadersHeight = 35
$dgv.DefaultCellStyle.BackColor = $BgDark
$dgv.DefaultCellStyle.ForeColor = $FgText
$dgv.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(60, 60, 90)
$dgv.DefaultCellStyle.SelectionForeColor = $Accent
$dgv.DefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(5, 3, 5, 3)
$dgv.RowHeadersVisible = $false
$dgv.AllowUserToAddRows = $false
$dgv.AllowUserToDeleteRows = $false
$dgv.ReadOnly = $true
$dgv.SelectionMode = "FullRowSelect"
$dgv.AutoSizeColumnsMode = "Fill"
$dgv.EnableHeadersVisualStyles = $false
$dgv.RowTemplate.Height = 30
$dgv.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(36, 36, 54)

$dgv.Columns.Add("Name", "Nome da GPO") | Out-Null
$dgv.Columns.Add("Status", "Status") | Out-Null
$dgv.Columns.Add("LinkedOUs", "Vinculada em") | Out-Null
$dgv.Columns.Add("Created", "Criada em") | Out-Null
$dgv.Columns.Add("Modified", "Modificada em") | Out-Null
$dgv.Columns.Add("Description", "Descricao") | Out-Null
$dgv.Columns.Add("Id", "ID") | Out-Null
$dgv.Columns["Name"].FillWeight = 22
$dgv.Columns["Status"].FillWeight = 13
$dgv.Columns["LinkedOUs"].FillWeight = 18
$dgv.Columns["Created"].FillWeight = 12
$dgv.Columns["Modified"].FillWeight = 12
$dgv.Columns["Description"].FillWeight = 30
$dgv.Columns["Id"].Visible = $false

# Label de estado vazio
$lblEmpty = New-Object System.Windows.Forms.Label
$lblEmpty.Text = "Nenhuma GPO encontrada.`nClique em 'Nova GPO' ou 'GPO Rapida' para comecar."
$lblEmpty.AutoSize = $false
$lblEmpty.Size = New-Object System.Drawing.Size(500, 80)
$lblEmpty.TextAlign = "MiddleCenter"
$lblEmpty.ForeColor = $Overlay
$lblEmpty.Font = New-Object System.Drawing.Font("Segoe UI", 12)
$lblEmpty.Anchor = "None"
$lblEmpty.Visible = $false
$tabGPOs.Controls.Add($lblEmpty)

$tabGPOs.Controls.Add($dgv)

# ══════════════════════════════════════════
#  ABA CADASTRO DE APPS (banco local JSON)
# ══════════════════════════════════════════
$script:AppsDbPath = Join-Path $PSScriptRoot "apps_db.json"

function Load-AppsDb {
    if (Test-Path $script:AppsDbPath) {
        try {
            $raw = Get-Content $script:AppsDbPath -Raw -Encoding UTF8
            $arr = $raw | ConvertFrom-Json
            return @($arr)
        } catch { return @() }
    }
    return @()
}

function Save-AppsDb {
    param($apps)
    $apps | ConvertTo-Json -Depth 5 | Set-Content $script:AppsDbPath -Force -Encoding UTF8
}

$script:AllApps = Load-AppsDb

# Layout da aba Apps
$pnlAppsTop = New-Object System.Windows.Forms.Panel
$pnlAppsTop.Dock = "Top"; $pnlAppsTop.Height = 55; $pnlAppsTop.BackColor = $BgPanel

$lblAppsTitle = New-Object System.Windows.Forms.Label
$lblAppsTitle.Text = "CADASTRO DE APLICATIVOS"
$lblAppsTitle.AutoSize = $true; $lblAppsTitle.ForeColor = [System.Drawing.Color]::FromArgb(203,166,247)
$lblAppsTitle.Font = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)
$lblAppsTitle.Location = New-Object System.Drawing.Point(15, 5)
$pnlAppsTop.Controls.Add($lblAppsTitle)

$lblAppsSub = New-Object System.Windows.Forms.Label
$lblAppsSub.Text = "Cadastre aqui os aplicativos. Na hora de criar/editar uma GPO, eles aparecem automaticamente."
$lblAppsSub.AutoSize = $true; $lblAppsSub.ForeColor = [System.Drawing.Color]::FromArgb(108,112,134)
$lblAppsSub.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblAppsSub.Location = New-Object System.Drawing.Point(15, 32)
$pnlAppsTop.Controls.Add($lblAppsSub)
$tabApps.Controls.Add($pnlAppsTop)

# Toolbar Apps
$pnlAppsTool = New-Object System.Windows.Forms.Panel
$pnlAppsTool.Dock = "Top"; $pnlAppsTool.Height = 48; $pnlAppsTool.BackColor = [System.Drawing.Color]::FromArgb(40,40,58)

$btnAddApp = New-Object System.Windows.Forms.Button
$btnAddApp.Text = "+ Novo App"; $btnAddApp.Location = New-Object System.Drawing.Point(10, 9)
$btnAddApp.Size = New-Object System.Drawing.Size(120, 30); $btnAddApp.BackColor = $Green; $btnAddApp.ForeColor = $BgDark
$btnAddApp.FlatStyle = "Flat"; $btnAddApp.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnAddApp.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$pnlAppsTool.Controls.Add($btnAddApp)

$btnEditApp = New-Object System.Windows.Forms.Button
$btnEditApp.Text = "Editar"; $btnEditApp.Location = New-Object System.Drawing.Point(140, 9)
$btnEditApp.Size = New-Object System.Drawing.Size(90, 30); $btnEditApp.BackColor = $Yellow; $btnEditApp.ForeColor = $BgDark
$btnEditApp.FlatStyle = "Flat"; $btnEditApp.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnEditApp.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$pnlAppsTool.Controls.Add($btnEditApp)

$btnDelApp = New-Object System.Windows.Forms.Button
$btnDelApp.Text = "Excluir"; $btnDelApp.Location = New-Object System.Drawing.Point(240, 9)
$btnDelApp.Size = New-Object System.Drawing.Size(90, 30); $btnDelApp.BackColor = $Red; $btnDelApp.ForeColor = [System.Drawing.Color]::White
$btnDelApp.FlatStyle = "Flat"; $btnDelApp.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnDelApp.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$pnlAppsTool.Controls.Add($btnDelApp)

$btnDupApp = New-Object System.Windows.Forms.Button
$btnDupApp.Text = "Duplicar"; $btnDupApp.Location = New-Object System.Drawing.Point(340, 9)
$btnDupApp.Size = New-Object System.Drawing.Size(90, 30); $btnDupApp.BackColor = $BgField; $btnDupApp.ForeColor = $FgText
$btnDupApp.FlatStyle = "Flat"; $btnDupApp.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnDupApp.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$pnlAppsTool.Controls.Add($btnDupApp)

$txtAppSearch = New-Object System.Windows.Forms.TextBox
$txtAppSearch.Location = New-Object System.Drawing.Point(460, 11); $txtAppSearch.Size = New-Object System.Drawing.Size(200, 26)
$txtAppSearch.BackColor = $BgField; $txtAppSearch.ForeColor = $FgText
$txtAppSearch.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$pnlAppsTool.Controls.Add($txtAppSearch)

$lblAppSearch = New-Object System.Windows.Forms.Label
$lblAppSearch.Text = "Buscar:"; $lblAppSearch.AutoSize = $true
$lblAppSearch.Location = New-Object System.Drawing.Point(460, -2); $lblAppSearch.ForeColor = [System.Drawing.Color]::FromArgb(108,112,134)
$lblAppSearch.Font = New-Object System.Drawing.Font("Segoe UI", 7)
$pnlAppsTool.Controls.Add($lblAppSearch)

$lblAppCount = New-Object System.Windows.Forms.Label
$lblAppCount.AutoSize = $true; $lblAppCount.ForeColor = $Yellow
$lblAppCount.Location = New-Object System.Drawing.Point(680, 15)
$pnlAppsTool.Controls.Add($lblAppCount)

$tabApps.Controls.Add($pnlAppsTool)

# DataGridView Apps
$dgvApps = New-Object System.Windows.Forms.DataGridView
$dgvApps.Dock = "Fill"; $dgvApps.BackgroundColor = $BgDark; $dgvApps.ForeColor = $FgText; $dgvApps.GridColor = $BgField
$dgvApps.BorderStyle = "None"; $dgvApps.CellBorderStyle = "SingleHorizontal"
$dgvApps.ColumnHeadersDefaultCellStyle.BackColor = $BgPanel
$dgvApps.ColumnHeadersDefaultCellStyle.ForeColor = $Accent
$dgvApps.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$dgvApps.ColumnHeadersHeight = 35
$dgvApps.DefaultCellStyle.BackColor = $BgDark; $dgvApps.DefaultCellStyle.ForeColor = $FgText
$dgvApps.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(60,60,90)
$dgvApps.DefaultCellStyle.SelectionForeColor = $Accent
$dgvApps.DefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(5, 3, 5, 3)
$dgvApps.RowHeadersVisible = $false; $dgvApps.AllowUserToAddRows = $false
$dgvApps.AllowUserToDeleteRows = $false; $dgvApps.ReadOnly = $true
$dgvApps.SelectionMode = "FullRowSelect"; $dgvApps.AutoSizeColumnsMode = "Fill"
$dgvApps.EnableHeadersVisualStyles = $false; $dgvApps.RowTemplate.Height = 30
$dgvApps.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(36, 36, 54)

$dgvApps.Columns.Add("AppName", "Nome") | Out-Null
$dgvApps.Columns.Add("AppCategory", "Categoria") | Out-Null
$dgvApps.Columns.Add("AppVersion", "Versao") | Out-Null
$dgvApps.Columns.Add("AppLink", "Link / Caminho") | Out-Null
$dgvApps.Columns.Add("AppSilent", "Instalacao Silenciosa") | Out-Null
$dgvApps.Columns.Add("AppDesc", "Observacoes") | Out-Null
$dgvApps.Columns["AppName"].FillWeight = 18
$dgvApps.Columns["AppCategory"].FillWeight = 12
$dgvApps.Columns["AppVersion"].FillWeight = 8
$dgvApps.Columns["AppLink"].FillWeight = 25
$dgvApps.Columns["AppSilent"].FillWeight = 17
$dgvApps.Columns["AppDesc"].FillWeight = 20

function Render-AppsTable {
    param([string]$filter)
    $dgvApps.Rows.Clear()
    foreach ($app in $script:AllApps) {
        if ($filter) {
            $match = $app.Name.ToLower().Contains($filter.ToLower()) -or $app.Category.ToLower().Contains($filter.ToLower()) -or $app.Desc.ToLower().Contains($filter.ToLower())
            if (-not $match) { continue }
        }
        $dgvApps.Rows.Add($app.Name, $app.Category, $app.Version, $app.Link, $app.Silent, $app.Desc) | Out-Null
    }
    $lblAppCount.Text = "$($dgvApps.Rows.Count) apps"
}

Render-AppsTable ""
$txtAppSearch.Add_TextChanged({ Render-AppsTable $txtAppSearch.Text })

# ── Formulario de app (Novo / Editar) ──
function Show-AppForm {
    param([hashtable]$existing)

    $af = New-Object System.Windows.Forms.Form
    $af.Text = if ($existing) { "Editar App" } else { "Novo App" }
    $af.Size = New-Object System.Drawing.Size(560, 520)
    $af.StartPosition = "CenterParent"; $af.BackColor = $BgDark; $af.ForeColor = $FgText
    $af.FormBorderStyle = "FixedDialog"; $af.MaximizeBox = $false; $af.MinimizeBox = $false
    $af.Font = New-Object System.Drawing.Font("Segoe UI", 10)

    $yy = 15
    # Nome
    $lN = New-Object System.Windows.Forms.Label; $lN.Text = "Nome do Aplicativo *"; $lN.AutoSize = $true
    $lN.Location = New-Object System.Drawing.Point(15, $yy); $lN.ForeColor = $Accent
    $lN.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $af.Controls.Add($lN); $yy += 22
    $tN = New-Object System.Windows.Forms.TextBox; $tN.Location = New-Object System.Drawing.Point(15, $yy)
    $tN.Size = New-Object System.Drawing.Size(510, 28); $tN.BackColor = $BgField; $tN.ForeColor = $FgText
    $tN.Font = New-Object System.Drawing.Font("Segoe UI", 12)
    if ($existing) { $tN.Text = $existing.Name }
    $af.Controls.Add($tN); $yy += 40

    # Categoria
    $lC = New-Object System.Windows.Forms.Label; $lC.Text = "Categoria"; $lC.AutoSize = $true
    $lC.Location = New-Object System.Drawing.Point(15, $yy); $lC.ForeColor = $Accent
    $lC.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $af.Controls.Add($lC); $yy += 22
    $tC = New-Object System.Windows.Forms.ComboBox; $tC.Location = New-Object System.Drawing.Point(15, $yy)
    $tC.Size = New-Object System.Drawing.Size(250, 28); $tC.BackColor = $BgField; $tC.ForeColor = $FgText
    $tC.DropDownStyle = "DropDown"
    $tC.Items.AddRange(@("Navegador","Utilitario","Seguranca","Comunicacao","Office","Desenvolvimento","Driver","Sistema","Acesso Remoto","Outro"))
    if ($existing) { $tC.Text = $existing.Category } else { $tC.SelectedIndex = 0 }
    $af.Controls.Add($tC); $yy += 40

    # Versao
    $lV = New-Object System.Windows.Forms.Label; $lV.Text = "Versao"; $lV.AutoSize = $true
    $lV.Location = New-Object System.Drawing.Point(15, $yy); $lV.ForeColor = $Accent
    $lV.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $af.Controls.Add($lV); $yy += 22
    $tV = New-Object System.Windows.Forms.TextBox; $tV.Location = New-Object System.Drawing.Point(15, $yy)
    $tV.Size = New-Object System.Drawing.Size(200, 28); $tV.BackColor = $BgField; $tV.ForeColor = $FgText
    if ($existing) { $tV.Text = $existing.Version }
    $af.Controls.Add($tV); $yy += 40

    # Link / Caminho do instalador
    $lL = New-Object System.Windows.Forms.Label; $lL.Text = "Link de Download ou Caminho de Rede"; $lL.AutoSize = $true
    $lL.Location = New-Object System.Drawing.Point(15, $yy); $lL.ForeColor = $Accent
    $lL.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $af.Controls.Add($lL); $yy += 22
    $tL = New-Object System.Windows.Forms.TextBox; $tL.Location = New-Object System.Drawing.Point(15, $yy)
    $tL.Size = New-Object System.Drawing.Size(430, 28); $tL.BackColor = $BgField; $tL.ForeColor = $FgText
    if ($existing) { $tL.Text = $existing.Link }
    $af.Controls.Add($tL)

    $btnBrowse = New-Object System.Windows.Forms.Button; $btnBrowse.Text = "..."
    $btnBrowse.Location = New-Object System.Drawing.Point(450, $yy); $btnBrowse.Size = New-Object System.Drawing.Size(40, 28)
    $btnBrowse.BackColor = $BgField; $btnBrowse.ForeColor = $FgText; $btnBrowse.FlatStyle = "Flat"
    $btnBrowse.Add_Click({
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Filter = "Instaladores (*.exe;*.msi;*.msp)|*.exe;*.msi;*.msp|Todos (*.*)|*.*"
        $ofd.Title = "Selecionar instalador"
        if ($ofd.ShowDialog() -eq "OK") { $tL.Text = $ofd.FileName }
    })
    $af.Controls.Add($btnBrowse)

    $btnPaste = New-Object System.Windows.Forms.Button; $btnPaste.Text = "Colar"
    $btnPaste.Location = New-Object System.Drawing.Point(495, $yy); $btnPaste.Size = New-Object System.Drawing.Size(30, 28)
    $btnPaste.BackColor = $BgField; $btnPaste.ForeColor = $FgText; $btnPaste.FlatStyle = "Flat"
    $btnPaste.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $btnPaste.Add_Click({ $tL.Text = [System.Windows.Forms.Clipboard]::GetText() })
    $af.Controls.Add($btnPaste)
    $yy += 40

    # Comando de instalacao silenciosa
    $lS = New-Object System.Windows.Forms.Label; $lS.Text = "Comando de Instalacao Silenciosa"; $lS.AutoSize = $true
    $lS.Location = New-Object System.Drawing.Point(15, $yy); $lS.ForeColor = $Accent
    $lS.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $af.Controls.Add($lS); $yy += 22
    $tS = New-Object System.Windows.Forms.TextBox; $tS.Location = New-Object System.Drawing.Point(15, $yy)
    $tS.Size = New-Object System.Drawing.Size(510, 28); $tS.BackColor = $BgField; $tS.ForeColor = $FgText
    $tS.Font = New-Object System.Drawing.Font("Consolas", 10)
    if ($existing) { $tS.Text = $existing.Silent } else { $tS.Text = "/S /silent /quiet /norestart" }
    $af.Controls.Add($tS); $yy += 40

    # Observacoes
    $lD = New-Object System.Windows.Forms.Label; $lD.Text = "Observacoes"; $lD.AutoSize = $true
    $lD.Location = New-Object System.Drawing.Point(15, $yy); $lD.ForeColor = $Accent
    $lD.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $af.Controls.Add($lD); $yy += 22
    $tD = New-Object System.Windows.Forms.TextBox; $tD.Location = New-Object System.Drawing.Point(15, $yy)
    $tD.Size = New-Object System.Drawing.Size(510, 50); $tD.BackColor = $BgField; $tD.ForeColor = $FgText
    $tD.Multiline = $true
    if ($existing) { $tD.Text = $existing.Desc }
    $af.Controls.Add($tD); $yy += 65

    # Botoes
    $bSave = New-Object System.Windows.Forms.Button; $bSave.Text = "SALVAR"
    $bSave.Location = New-Object System.Drawing.Point(150, $yy); $bSave.Size = New-Object System.Drawing.Size(120, 36)
    $bSave.BackColor = $Green; $bSave.ForeColor = $BgDark; $bSave.FlatStyle = "Flat"
    $bSave.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $bSave.DialogResult = "OK"
    $af.Controls.Add($bSave)
    $af.AcceptButton = $bSave

    $bCancel = New-Object System.Windows.Forms.Button; $bCancel.Text = "Cancelar"
    $bCancel.Location = New-Object System.Drawing.Point(290, $yy); $bCancel.Size = New-Object System.Drawing.Size(100, 36)
    $bCancel.BackColor = $BgField; $bCancel.ForeColor = $Red; $bCancel.FlatStyle = "Flat"
    $bCancel.DialogResult = "Cancel"
    $af.Controls.Add($bCancel)

    $af.Tag = @{ tN=$tN; tC=$tC; tV=$tV; tL=$tL; tS=$tS; tD=$tD }
    $result = $af.ShowDialog()

    if ($result -eq "OK" -and -not [string]::IsNullOrEmpty($tN.Text.Trim())) {
        $app = @{
            Name     = $tN.Text.Trim()
            Category = $tC.Text.Trim()
            Version  = $tV.Text.Trim()
            Link     = $tL.Text.Trim()
            Silent   = $tS.Text.Trim()
            Desc     = $tD.Text.Trim()
            Id       = if ($existing) { $existing.Id } else { [Guid]::NewGuid().ToString() }
        }
        $af.Dispose()
        return $app
    }
    $af.Dispose()
    return $null
}

# ── Eventos dos botoes de Apps ──
$btnAddApp.Add_Click({
    $newApp = Show-AppForm
    if ($newApp) {
        $script:AllApps += $newApp
        Save-AppsDb $script:AllApps
        Render-AppsTable $txtAppSearch.Text
    }
})

$btnEditApp.Add_Click({
    if ($dgvApps.SelectedRows.Count -eq 0) {
        Show-DarkMsg "Selecione um app para editar." "Aviso" "OK" "Warning"; return
    }
    $idx = $dgvApps.SelectedRows[0].Index
    $appName = $dgvApps.SelectedRows[0].Cells["AppName"].Value
    # Encontrar no array
    $existing = $script:AllApps | Where-Object { $_.Name -eq $appName } | Select-Object -First 1
    if (-not $existing) { return }
    $existHash = @{
        Name = $existing.Name; Category = $existing.Category; Version = $existing.Version
        Link = $existing.Link; Silent = $existing.Silent; Desc = $existing.Desc
        Id = if ($existing.Id) { $existing.Id } else { [Guid]::NewGuid().ToString() }
    }
    $edited = Show-AppForm -existing $existHash
    if ($edited) {
        # Substituir
        $newList = @()
        $replaced = $false
        foreach ($a in $script:AllApps) {
            if ($a.Name -eq $appName -and -not $replaced) { $newList += $edited; $replaced = $true }
            else { $newList += $a }
        }
        $script:AllApps = $newList
        Save-AppsDb $script:AllApps
        Render-AppsTable $txtAppSearch.Text
    }
})

$btnDelApp.Add_Click({
    if ($dgvApps.SelectedRows.Count -eq 0) {
        Show-DarkMsg "Selecione um app para excluir." "Aviso" "OK" "Warning"; return
    }
    $appName = $dgvApps.SelectedRows[0].Cells["AppName"].Value
    $conf = Show-DarkMsg "Excluir o app '$appName'?" "Confirmar" "YesNo" "Warning"
    if ($conf -eq "Yes") {
        $script:AllApps = @($script:AllApps | Where-Object { $_.Name -ne $appName })
        Save-AppsDb $script:AllApps
        Render-AppsTable $txtAppSearch.Text
    }
})

$btnDupApp.Add_Click({
    if ($dgvApps.SelectedRows.Count -eq 0) {
        Show-DarkMsg "Selecione um app para duplicar." "Aviso" "OK" "Warning"; return
    }
    $appName = $dgvApps.SelectedRows[0].Cells["AppName"].Value
    $existing = $script:AllApps | Where-Object { $_.Name -eq $appName } | Select-Object -First 1
    if ($existing) {
        $dup = @{
            Name = "$($existing.Name) (copia)"; Category = $existing.Category; Version = $existing.Version
            Link = $existing.Link; Silent = $existing.Silent; Desc = $existing.Desc
            Id = [Guid]::NewGuid().ToString()
        }
        $script:AllApps += $dup
        Save-AppsDb $script:AllApps
        Render-AppsTable $txtAppSearch.Text
    }
})

# Duplo-clique edita
$dgvApps.Add_CellDoubleClick({
    $btnEditApp.PerformClick()
})

$tabApps.Controls.Add($dgvApps)

# Montar TabControl principal
$mainTabs.TabPages.Add($tabGPOs)
$mainTabs.TabPages.Add($tabApps)

# Status bar
$statusBar = New-Object System.Windows.Forms.StatusStrip
$statusBar.BackColor = $BgPanel
$statusBar.SizingGrip = $false
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.ForeColor = $FgText
$statusLabel.Text = "Pronto"
$statusLabel.Spring = $true
$statusLabel.TextAlign = "MiddleLeft"
$statusBar.Items.Add($statusLabel) | Out-Null

$statusTime = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusTime.ForeColor = $Overlay
$statusTime.Text = ""
$statusTime.Alignment = "Right"
$statusBar.Items.Add($statusTime) | Out-Null

# Timer para relogio na statusbar
$clockTimer = New-Object System.Windows.Forms.Timer
$clockTimer.Interval = 1000
$clockTimer.Add_Tick({ $statusTime.Text = (Get-Date).ToString("dd/MM/yyyy HH:mm:ss") })
$clockTimer.Start()

$dashPanel.Controls.Add($mainTabs)
$dashPanel.Controls.Add($statusBar)

$form.Controls.Add($dashPanel)

# ══════════════════════════════════════════
#  FUNCOES
# ══════════════════════════════════════════

# Armazena todas as GPOs para filtro
$script:AllGPOs = @()

function Do-Login {
    $dom  = $txtDomain.Text.Trim()
    $user = $txtUser.Text.Trim()
    $pass = $txtPass.Text

    if ([string]::IsNullOrEmpty($dom)) {
        $lblMsg.ForeColor = $Red; $lblMsg.Text = "Informe o dominio."; return
    }
    if ([string]::IsNullOrEmpty($user)) {
        $lblMsg.ForeColor = $Red; $lblMsg.Text = "Informe o usuario."; return
    }
    if ([string]::IsNullOrEmpty($pass)) {
        $lblMsg.ForeColor = $Red; $lblMsg.Text = "Informe a senha."; return
    }

    $lblMsg.ForeColor = $Yellow
    $lblMsg.Text = "Autenticando em $dom..."
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    [System.Windows.Forms.Application]::DoEvents()

    try {
        # Autenticar via LDAP DirectoryEntry (mais confiavel que PrincipalContext)
        $ldapPath = "LDAP://$dom"
        $entry = New-Object System.DirectoryServices.DirectoryEntry($ldapPath, "$dom\$user", $pass)
        
        # Forcar bind - se falhar, credenciais invalidas
        $null = $entry.distinguishedName
        
        if ($entry.distinguishedName) {
            $valid = $true
        } else {
            # Fallback: tentar com user@domain
            $entry2 = New-Object System.DirectoryServices.DirectoryEntry($ldapPath, "$user@$dom", $pass)
            $null = $entry2.distinguishedName
            $valid = [bool]$entry2.distinguishedName
            if ($entry2) { $entry2.Dispose() }
        }
        
        if ($entry) { $entry.Dispose() }

        if ($valid) {
            $secPass = ConvertTo-SecureString $pass -AsPlainText -Force
            $script:Credential = New-Object System.Management.Automation.PSCredential("$dom\$user", $secPass)
            $script:Domain = $dom

            $lblInfo.Text = "$user @ $dom"
            $loginPanel.Visible = $false
            $dashPanel.Visible = $true

            # Posicionar botao sair
            $btnLogout.Location = New-Object System.Drawing.Point(($header.ClientSize.Width - 80), 12)

            Load-GPOs
        } else {
            $lblMsg.ForeColor = $Red
            $lblMsg.Text = "Credenciais invalidas!"
            $txtPass.Text = ""
            $txtPass.Focus()
        }
    } catch {
        $errMsg = $_.Exception.Message
        if ($_.Exception.InnerException) { $errMsg = $_.Exception.InnerException.Message }
        
        # Erro de credencial vem como "Logon failure" ou similar
        if ($errMsg -match "logon|password|credential|senha|acesso negado|access denied") {
            $lblMsg.ForeColor = $Red
            $lblMsg.Text = "Credenciais invalidas!"
            $txtPass.Text = ""
            $txtPass.Focus()
        } else {
            $lblMsg.ForeColor = $Red
            $lblMsg.Text = "Erro: $errMsg"
        }
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
}

function Do-Logout {
    $script:Credential = $null
    $dashPanel.Visible = $false
    $loginPanel.Visible = $true
    $txtPass.Text = ""
    $lblMsg.Text = ""
    $dgv.Rows.Clear()
    $script:AllGPOs = @()
    Center-LoginPanel
    $txtPass.Focus()
}

function Load-GPOs {
    $lblGPOCount.ForeColor = $Yellow
    $lblGPOCount.Text = "Carregando..."
    $statusLabel.Text = "Carregando GPOs do dominio $($script:Domain)..."
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    [System.Windows.Forms.Application]::DoEvents()

    try {
        # Buscar GPOs via LDAP direto (nao precisa de modulo GroupPolicy)
        $domDN = ($script:Domain.Split('.') | ForEach-Object { "DC=$_" }) -join ','
        $searcher = New-Object System.DirectoryServices.DirectorySearcher
        $searcher.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://CN=Policies,CN=System,$domDN")
        $searcher.Filter = "(objectClass=groupPolicyContainer)"
        $searcher.PageSize = 1000
        $searcher.PropertiesToLoad.AddRange(@("displayName","cn","gPCFileSysPath","whenCreated","whenChanged","flags","gPCUserExtensionNames","gPCMachineExtensionNames"))
        
        $results = $searcher.FindAll()

        # Buscar todos links de GPO em OUs/containers/dominio para mostrar "Vinculada em"
        $ouLinks = @{}
        try {
            $sOU = New-Object System.DirectoryServices.DirectorySearcher
            $sOU.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$domDN")
            $sOU.Filter = "(gPLink=*)"
            $sOU.PageSize = 1000
            $sOU.SearchScope = "Subtree"
            $sOU.PropertiesToLoad.AddRange(@("distinguishedName","name","gPLink")) | Out-Null
            $ouResults = $sOU.FindAll()
            foreach ($ouR in $ouResults) {
                $ouName = if ($ouR.Properties["name"].Count -gt 0) { [string]$ouR.Properties["name"][0] } else { "" }
                $gpLink = if ($ouR.Properties["gplink"].Count -gt 0) { [string]$ouR.Properties["gplink"][0] } else { "" }
                if ($gpLink) {
                    # Extrair GUIDs do gPLink
                    $matches2 = [regex]::Matches($gpLink, '\{([0-9A-Fa-f\-]+)\}')
                    foreach ($m in $matches2) {
                        $guid = "{" + $m.Groups[1].Value.ToUpper() + "}"
                        if (-not $ouLinks.ContainsKey($guid)) { $ouLinks[$guid] = [System.Collections.ArrayList]@() }
                        $ouLinks[$guid].Add($ouName) | Out-Null
                    }
                }
            }
            $ouResults.Dispose(); $sOU.Dispose()
        } catch {}

        $script:AllGPOs = @()
        foreach ($r in $results) {
            $props = $r.Properties
            $name = if ($props["displayname"].Count -gt 0) { $props["displayname"][0] } else { "Sem nome" }
            $guid = if ($props["cn"].Count -gt 0) { $props["cn"][0] } else { "" }
            $created = if ($props["whencreated"].Count -gt 0) { ([datetime]$props["whencreated"][0]).ToString("yyyy-MM-dd HH:mm") } else { "-" }
            $modified = if ($props["whenchanged"].Count -gt 0) { ([datetime]$props["whenchanged"][0]).ToString("yyyy-MM-dd HH:mm") } else { "-" }
            $flags = if ($props["flags"].Count -gt 0) { [int]$props["flags"][0] } else { 0 }
            
            # flags: 0=AllEnabled, 1=UserDisabled, 2=ComputerDisabled, 3=AllDisabled
            $status = switch ($flags) {
                0 { "Habilitada" }
                1 { "Config Usuario Desabilitada" }
                2 { "Config Computador Desabilitada" }
                3 { "Toda Desabilitada" }
                default { "Desconhecido ($flags)" }
            }

            $script:AllGPOs += [PSCustomObject]@{
                Name        = $name
                Status      = $status
                Created     = $created
                Modified    = $modified
                Description = ""
                Id          = $guid
                LinkedOUs   = if ($ouLinks.ContainsKey($guid.ToUpper())) { ($ouLinks[$guid.ToUpper()] -join ", ") } else { "" }
            }
        }
        
        $results.Dispose()
        $searcher.Dispose()

        # Ordenar por nome
        $script:AllGPOs = $script:AllGPOs | Sort-Object Name

        Render-GPOTable $script:AllGPOs

        $lblGPOCount.ForeColor = $Green
        $lblGPOCount.Text = "$($script:AllGPOs.Count) GPOs"
        $statusLabel.Text = "GPOs carregadas com sucesso via LDAP."
    } catch {
        $lblGPOCount.ForeColor = $Red
        $lblGPOCount.Text = "Erro ao carregar"
        $statusLabel.Text = "Erro: $($_.Exception.Message)"
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
}

function Render-GPOTable {
    param($gpoList)
    $dgv.Rows.Clear()
    if (-not $gpoList -or @($gpoList).Count -eq 0) {
        $lblEmpty.Visible = $true
        $lblEmpty.Location = New-Object System.Drawing.Point(
            [int](($tabGPOs.ClientSize.Width - $lblEmpty.Width) / 2),
            [int](($tabGPOs.ClientSize.Height - $lblEmpty.Height) / 2)
        )
        return
    }
    $lblEmpty.Visible = $false
    foreach ($g in $gpoList) {
        $linkedText = if ($g.LinkedOUs) { $g.LinkedOUs } else { "-" }
        $row = $dgv.Rows.Add($g.Name, $g.Status, $linkedText, $g.Created, $g.Modified, $g.Description, $g.Id)
        $cell = $dgv.Rows[$row].Cells["Status"]
        if ($g.Status -eq "Habilitada") {
            $cell.Style.ForeColor = $Green
        } elseif ($g.Status -eq "Toda Desabilitada") {
            $cell.Style.ForeColor = $Red
        } else {
            $cell.Style.ForeColor = $Yellow
        }
        # Colorir linked OUs
        $linkedCell = $dgv.Rows[$row].Cells["LinkedOUs"]
        if ($linkedText -ne "-") { $linkedCell.Style.ForeColor = $Accent }
    }
}

function Filter-GPOs {
    $term = $txtSearch.Text.Trim().ToLower()
    if ([string]::IsNullOrEmpty($term)) {
        Render-GPOTable $script:AllGPOs
        $lblGPOCount.Text = "$($script:AllGPOs.Count) GPOs"
    } else {
        $filtered = $script:AllGPOs | Where-Object {
            $_.Name.ToLower().Contains($term) -or $_.Description.ToLower().Contains($term) -or $_.LinkedOUs.ToLower().Contains($term)
        }
        Render-GPOTable $filtered
        $lblGPOCount.Text = "$(@($filtered).Count) / $($script:AllGPOs.Count) GPOs"
    }
}

function Create-NewGPO {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Nova GPO"
    $dlg.Size = New-Object System.Drawing.Size(420, 230)
    $dlg.StartPosition = "CenterParent"
    $dlg.BackColor = $BgPanel
    $dlg.ForeColor = $FgText
    $dlg.FormBorderStyle = "FixedDialog"
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false

    $l1 = New-Object System.Windows.Forms.Label
    $l1.Text = "Nome da GPO:"; $l1.AutoSize = $true; $l1.Location = New-Object System.Drawing.Point(15, 15)
    $dlg.Controls.Add($l1)

    $tName = New-Object System.Windows.Forms.TextBox
    $tName.Location = New-Object System.Drawing.Point(15, 38); $tName.Size = New-Object System.Drawing.Size(370, 26)
    $tName.BackColor = $BgField; $tName.ForeColor = $FgText
    $dlg.Controls.Add($tName)

    $l2 = New-Object System.Windows.Forms.Label
    $l2.Text = "Comentario (opcional):"; $l2.AutoSize = $true; $l2.Location = New-Object System.Drawing.Point(15, 75)
    $dlg.Controls.Add($l2)

    $tComment = New-Object System.Windows.Forms.TextBox
    $tComment.Location = New-Object System.Drawing.Point(15, 98); $tComment.Size = New-Object System.Drawing.Size(370, 26)
    $tComment.BackColor = $BgField; $tComment.ForeColor = $FgText
    $dlg.Controls.Add($tComment)

    $bOk = New-Object System.Windows.Forms.Button
    $bOk.Text = "Criar"; $bOk.Location = New-Object System.Drawing.Point(145, 145); $bOk.Size = New-Object System.Drawing.Size(120, 34)
    $bOk.BackColor = $Green; $bOk.ForeColor = $BgDark; $bOk.FlatStyle = "Flat"; $bOk.DialogResult = "OK"
    $dlg.Controls.Add($bOk)
    $dlg.AcceptButton = $bOk

    if ($dlg.ShowDialog() -eq "OK" -and -not [string]::IsNullOrEmpty($tName.Text.Trim())) {
        try {
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            # Criar GPO via ADSI
            $domDN = ($script:Domain.Split('.') | ForEach-Object { "DC=$_" }) -join ','
            $policiesPath = "LDAP://CN=Policies,CN=System,$domDN"
            $policies = [ADSI]$policiesPath
            $newGuid = [Guid]::NewGuid().ToString('B').ToUpper()
            $newGPO = $policies.Create("groupPolicyContainer", "CN=$newGuid")
            $newGPO.Put("displayName", $tName.Text.Trim())
            $newGPO.Put("gPCFunctionalityVersion", 2)
            $newGPO.Put("flags", 0)
            $newGPO.Put("versionNumber", 0)
            $gpcPath = "\\$($script:Domain)\SysVol\$($script:Domain)\Policies\$newGuid"
            $newGPO.Put("gPCFileSysPath", $gpcPath)
            $newGPO.SetInfo()
            
            # Criar pasta SYSVOL
            New-Item -Path $gpcPath -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
            New-Item -Path "$gpcPath\Machine" -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
            New-Item -Path "$gpcPath\User" -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null

            Show-DarkMsg "GPO '$($tName.Text)' criada!" "Sucesso" "OK" "Information"
            Load-GPOs
        } catch {
            Show-DarkMsg "Erro: $($_.Exception.Message)" "Falha" "OK" "Error"
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }
    $dlg.Dispose()
}

# ══════════════════════════════════════════
#  ASSISTENTE GPO RAPIDA (Wizard)
# ══════════════════════════════════════════
function New-QuickGPO {
    $domDN = ($script:Domain.Split('.') | ForEach-Object { "DC=$_" }) -join ','

    $wiz = New-Object System.Windows.Forms.Form
    $wiz.Text = "Assistente GPO Rapida"
    $wiz.Size = New-Object System.Drawing.Size(820, 620)
    $wiz.StartPosition = "CenterScreen"
    $wiz.BackColor = $BgDark; $wiz.ForeColor = $FgText
    $wiz.FormBorderStyle = "FixedDialog"
    $wiz.MaximizeBox = $false; $wiz.MinimizeBox = $false
    $wiz.Font = New-Object System.Drawing.Font("Segoe UI", 10)

    # ── Painel de steps (esquerda) ──
    $pnlSteps = New-Object System.Windows.Forms.Panel
    $pnlSteps.Dock = "Left"; $pnlSteps.Width = 180; $pnlSteps.BackColor = $BgPanel
    $pnlSteps.Padding = New-Object System.Windows.Forms.Padding(10, 20, 10, 10)

    $lblWizTitle = New-Object System.Windows.Forms.Label
    $lblWizTitle.Text = "GPO RAPIDA"; $lblWizTitle.AutoSize = $true
    $lblWizTitle.ForeColor = [System.Drawing.Color]::FromArgb(203,166,247)
    $lblWizTitle.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $lblWizTitle.Location = New-Object System.Drawing.Point(15, 15)
    $pnlSteps.Controls.Add($lblWizTitle)

    $stepLabels = @()
    $stepNames = @("1. Nome", "2. Politicas", "3. Destino", "4. Pronto!")
    for ($i = 0; $i -lt $stepNames.Count; $i++) {
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = $stepNames[$i]; $lbl.AutoSize = $true
        $lbl.Location = New-Object System.Drawing.Point(20, 70 + ($i * 40))
        $lbl.Font = New-Object System.Drawing.Font("Segoe UI", 11)
        $lbl.ForeColor = [System.Drawing.Color]::FromArgb(108,112,134)
        $pnlSteps.Controls.Add($lbl)
        $stepLabels += $lbl
    }
    $wiz.Controls.Add($pnlSteps)

    # ── Paineis de cada step ──
    $panels = @()

    # ═══ STEP 1: NOME ═══
    $step1 = New-Object System.Windows.Forms.Panel; $step1.Dock = "Fill"; $step1.Visible = $true
    $step1.Padding = New-Object System.Windows.Forms.Padding(30)

    $l1t = New-Object System.Windows.Forms.Label; $l1t.Text = "Como vai se chamar essa GPO?"
    $l1t.AutoSize = $true; $l1t.Location = New-Object System.Drawing.Point(30, 30)
    $l1t.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold); $l1t.ForeColor = $Accent
    $step1.Controls.Add($l1t)

    $l1s = New-Object System.Windows.Forms.Label; $l1s.Text = "Escolha um nome claro. Ex: Bloquear USB, Instalar Firefox, Papel de Parede Empresa"
    $l1s.AutoSize = $true; $l1s.Location = New-Object System.Drawing.Point(30, 70)
    $l1s.ForeColor = [System.Drawing.Color]::FromArgb(108,112,134)
    $step1.Controls.Add($l1s)

    $txtWizName = New-Object System.Windows.Forms.TextBox
    $txtWizName.Location = New-Object System.Drawing.Point(30, 110); $txtWizName.Size = New-Object System.Drawing.Size(520, 36)
    $txtWizName.BackColor = $BgField; $txtWizName.ForeColor = $FgText
    $txtWizName.Font = New-Object System.Drawing.Font("Segoe UI", 14)
    $step1.Controls.Add($txtWizName)

    $l1d = New-Object System.Windows.Forms.Label; $l1d.Text = "Descricao (opcional):"
    $l1d.AutoSize = $true; $l1d.Location = New-Object System.Drawing.Point(30, 165); $l1d.ForeColor = $FgText
    $step1.Controls.Add($l1d)

    $txtWizDesc = New-Object System.Windows.Forms.TextBox
    $txtWizDesc.Location = New-Object System.Drawing.Point(30, 190); $txtWizDesc.Size = New-Object System.Drawing.Size(520, 60)
    $txtWizDesc.BackColor = $BgField; $txtWizDesc.ForeColor = $FgText; $txtWizDesc.Multiline = $true
    $step1.Controls.Add($txtWizDesc)

    $panels += $step1

    # ═══ STEP 2: POLITICAS (checkboxes simples) ═══
    $step2 = New-Object System.Windows.Forms.Panel; $step2.Dock = "Fill"; $step2.Visible = $false

    $l2t = New-Object System.Windows.Forms.Label; $l2t.Text = "O que essa GPO vai fazer?"
    $l2t.AutoSize = $true; $l2t.Location = New-Object System.Drawing.Point(30, 15)
    $l2t.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold); $l2t.ForeColor = $Accent
    $step2.Controls.Add($l2t)

    $l2s = New-Object System.Windows.Forms.Label; $l2s.Text = "Marque tudo que quiser. Pode editar depois com mais detalhes."
    $l2s.AutoSize = $true; $l2s.Location = New-Object System.Drawing.Point(30, 48)
    $l2s.ForeColor = [System.Drawing.Color]::FromArgb(108,112,134)
    $step2.Controls.Add($l2s)

    # ScrollPanel para os checkboxes
    $scrollPol = New-Object System.Windows.Forms.Panel
    $scrollPol.Location = New-Object System.Drawing.Point(20, 75)
    $scrollPol.Size = New-Object System.Drawing.Size(580, 420)
    $scrollPol.AutoScroll = $true; $scrollPol.BackColor = $BgDark

    $wizPolicies = [ordered]@{
        "SEGURANCA"  = @(
            ,@("wizBlkUSB",      "Bloquear USB / Pen Drive")
            ,@("wizBlkCmd",      "Bloquear Prompt de Comando (CMD)")
            ,@("wizBlkRegedit",  "Bloquear Editor de Registro (regedit)")
            ,@("wizBlkPanel",    "Bloquear Painel de Controle")
            ,@("wizBlkTaskMgr",  "Bloquear Gerenciador de Tarefas")
            ,@("wizPwdPolicy",   "Exigir senha complexa (8+ chars)")
        )
        "APARENCIA"  = @(
            ,@("wizWallpaper",   "Definir papel de parede padrao")
            ,@("wizLockscreen",  "Definir tela de bloqueio padrao")
            ,@("wizNoDesktop",   "Ocultar icones da area de trabalho")
            ,@("wizNoTheme",     "Impedir troca de tema/papel parede")
        )
        "WINDOWS"    = @(
            ,@("wizNoUpdate",    "Desativar Windows Update automatico")
            ,@("wizNoStore",     "Bloquear Microsoft Store")
            ,@("wizNoCortana",   "Desativar Cortana")
            ,@("wizNoOneDrive",  "Desativar OneDrive")
            ,@("wizNoAutoPlay",  "Desativar AutoPlay de midias")
            ,@("wizNoTelemetry", "Desativar telemetria / privacidade")
        )
        "REDE"       = @(
            ,@("wizMapDrive",    "Mapear unidade de rede")
            ,@("wizProxy",       "Configurar proxy")
            ,@("wizFirewall",    "Ativar Firewall do Windows")
            ,@("wizNoWifi",      "Bloquear Wi-Fi")
        )
        "ENERGIA"    = @(
            ,@("wizNeverSleep",  "Nunca suspender / hibernar")
            ,@("wizScreenOff",   "Desligar tela apos 15 min")
        )
        "APPS"       = @(
            ,@("wizInstallApp",  "Instalar software (via script)")
            ,@("wizBlockApps",   "Bloquear aplicativos especificos")
        )
    }

    $script:wizChecks = @{}
    $yPos = 5
    foreach ($cat in $wizPolicies.Keys) {
        $lblCat = New-Object System.Windows.Forms.Label
        $lblCat.Text = $cat; $lblCat.AutoSize = $true
        $lblCat.Location = New-Object System.Drawing.Point(10, $yPos)
        $lblCat.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $lblCat.ForeColor = $Yellow
        $scrollPol.Controls.Add($lblCat)
        $yPos += 25

        foreach ($pol in $wizPolicies[$cat]) {
            $chk = New-Object System.Windows.Forms.CheckBox
            $chk.Text = $pol[1]
            $chk.AutoSize = $false
            $chk.Size = New-Object System.Drawing.Size(530, 24)
            $chk.Location = New-Object System.Drawing.Point(25, $yPos)
            $chk.ForeColor = $FgText; $chk.Font = New-Object System.Drawing.Font("Segoe UI", 10)
            $scrollPol.Controls.Add($chk)
            $script:wizChecks[$pol[0]] = $chk
            $yPos += 28
        }
        $yPos += 10
    }

    $step2.Controls.Add($scrollPol)
    $panels += $step2

    # ═══ STEP 3: DESTINO (TreeView com checkboxes) ═══
    $step3 = New-Object System.Windows.Forms.Panel; $step3.Dock = "Fill"; $step3.Visible = $false

    $l3t = New-Object System.Windows.Forms.Label; $l3t.Text = "Onde aplicar essa GPO?"
    $l3t.AutoSize = $true; $l3t.Location = New-Object System.Drawing.Point(30, 15)
    $l3t.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold); $l3t.ForeColor = $Accent
    $step3.Controls.Add($l3t)

    $l3s = New-Object System.Windows.Forms.Label; $l3s.Text = "Marque as pastas (OUs) onde a GPO sera aplicada. Todos dentro vao receber."
    $l3s.AutoSize = $true; $l3s.Location = New-Object System.Drawing.Point(30, 48)
    $l3s.ForeColor = [System.Drawing.Color]::FromArgb(108,112,134)
    $step3.Controls.Add($l3s)

    $wizTree = New-Object System.Windows.Forms.TreeView
    $wizTree.Location = New-Object System.Drawing.Point(25, 80)
    $wizTree.Size = New-Object System.Drawing.Size(570, 400)
    $wizTree.CheckBoxes = $true; $wizTree.BackColor = $BgDark; $wizTree.ForeColor = $FgText
    $wizTree.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $wizTree.LineColor = $BgField; $wizTree.BorderStyle = "FixedSingle"
    $wizTree.ShowPlusMinus = $true; $wizTree.ShowLines = $true; $wizTree.ShowRootLines = $true
    $wizTree.ItemHeight = 26

    # Construir arvore do AD (somente OUs - simplificado)
    function Build-WizTree {
        $wizTree.BeginUpdate(); $wizTree.Nodes.Clear()
        $rootNode = New-Object System.Windows.Forms.TreeNode
        $rootNode.Text = "$($script:Domain)"; $rootNode.Tag = $domDN
        $rootNode.NodeFont = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        $wizTree.Nodes.Add($rootNode) | Out-Null

        $nodeMap = @{}; $nodeMap[$domDN] = $rootNode

        try {
            $s = New-Object System.DirectoryServices.DirectorySearcher
            $s.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$domDN")
            $s.Filter = "(objectClass=organizationalUnit)"; $s.PageSize = 1000; $s.SearchScope = "Subtree"
            $s.PropertiesToLoad.AddRange(@("distinguishedName","name")) | Out-Null
            $res = $s.FindAll()
            $ous = @()
            foreach ($r in $res) {
                $ous += @{ DN = [string]$r.Properties["distinguishedname"][0]; Name = [string]$r.Properties["name"][0] }
            }
            $res.Dispose(); $s.Dispose()
            $ous = $ous | Sort-Object { $_.DN.Split(',').Count }

            foreach ($ou in $ous) {
                $parentDN = ($ou.DN -split ',', 2)[1]
                $parentNode = if ($nodeMap.ContainsKey($parentDN)) { $nodeMap[$parentDN] } else { $rootNode }
                $n = New-Object System.Windows.Forms.TreeNode
                $n.Text = $ou.Name; $n.Tag = $ou.DN
                $n.NodeFont = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
                $parentNode.Nodes.Add($n) | Out-Null
                $nodeMap[$ou.DN] = $n
            }
        } catch {}

        # Adicionar containers padroes
        try {
            $s = New-Object System.DirectoryServices.DirectorySearcher
            $s.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$domDN")
            $s.Filter = "(&(objectClass=container)(|(cn=Users)(cn=Computers)(cn=Builtin)))"; $s.PageSize = 10
            $s.PropertiesToLoad.AddRange(@("distinguishedName","cn")) | Out-Null
            $res = $s.FindAll()
            foreach ($r in $res) {
                $dn = [string]$r.Properties["distinguishedname"][0]
                $cn = [string]$r.Properties["cn"][0]
                $n = New-Object System.Windows.Forms.TreeNode
                $n.Text = $cn; $n.Tag = $dn
                $n.ForeColor = [System.Drawing.Color]::FromArgb(108,112,134)
                $rootNode.Nodes.Add($n) | Out-Null
                $nodeMap[$dn] = $n
            }
            $res.Dispose(); $s.Dispose()
        } catch {}

        # Mostrar quantidade de objetos em cada OU
        try {
            $s = New-Object System.DirectoryServices.DirectorySearcher
            $s.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$domDN")
            $s.Filter = "(|(objectClass=user)(objectClass=computer))"; $s.PageSize = 2000
            $s.PropertiesToLoad.Add("distinguishedName") | Out-Null
            $res = $s.FindAll()
            $countMap = @{}
            foreach ($r in $res) {
                $dn = [string]$r.Properties["distinguishedname"][0]
                $parent = ($dn -split ',', 2)[1]
                if (-not $countMap.ContainsKey($parent)) { $countMap[$parent] = 0 }
                $countMap[$parent]++
            }
            $res.Dispose(); $s.Dispose()

            foreach ($key in $nodeMap.Keys) {
                if ($countMap.ContainsKey($key)) {
                    $nodeMap[$key].Text = "$($nodeMap[$key].Text)  ($($countMap[$key]) objetos)"
                    $nodeMap[$key].ForeColor = $FgText
                }
            }
        } catch {}

        $rootNode.Expand()
        # expandir primeiro nivel
        foreach ($child in $rootNode.Nodes) { $child.Expand() }
        $wizTree.EndUpdate()
    }

    Build-WizTree

    function Get-CheckedWizOUs { param($nodes)
        $result = @()
        foreach ($n in $nodes) {
            if ($n.Checked -and $n.Tag) { $result += $n.Tag }
            if ($n.Nodes.Count -gt 0) { $result += Get-CheckedWizOUs $n.Nodes }
        }
        return $result
    }

    $step3.Controls.Add($wizTree)
    $panels += $step3

    # ═══ STEP 4: RESUMO ═══
    $step4 = New-Object System.Windows.Forms.Panel; $step4.Dock = "Fill"; $step4.Visible = $false

    $l4t = New-Object System.Windows.Forms.Label; $l4t.Text = "Tudo certo! Confirme e crie."
    $l4t.AutoSize = $true; $l4t.Location = New-Object System.Drawing.Point(30, 20)
    $l4t.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold); $l4t.ForeColor = $Green
    $step4.Controls.Add($l4t)

    $lblResume = New-Object System.Windows.Forms.Label
    $lblResume.Location = New-Object System.Drawing.Point(30, 65); $lblResume.Size = New-Object System.Drawing.Size(560, 400)
    $lblResume.ForeColor = $FgText; $lblResume.Font = New-Object System.Drawing.Font("Consolas", 10)
    $step4.Controls.Add($lblResume)

    $panels += $step4

    # Adicionar paineis ao form
    foreach ($p in $panels) { $wiz.Controls.Add($p) }

    # ── Barra inferior com botoes ──
    $pnlBottom = New-Object System.Windows.Forms.Panel
    $pnlBottom.Dock = "Bottom"; $pnlBottom.Height = 55; $pnlBottom.BackColor = $BgPanel

    $btnPrev = New-Object System.Windows.Forms.Button
    $btnPrev.Text = "< Voltar"; $btnPrev.Location = New-Object System.Drawing.Point(200, 12)
    $btnPrev.Size = New-Object System.Drawing.Size(120, 34); $btnPrev.BackColor = $BgField; $btnPrev.ForeColor = $FgText
    $btnPrev.FlatStyle = "Flat"; $btnPrev.Cursor = [System.Windows.Forms.Cursors]::Hand; $btnPrev.Enabled = $false
    $pnlBottom.Controls.Add($btnPrev)

    $btnNext = New-Object System.Windows.Forms.Button
    $btnNext.Text = "Proximo >"; $btnNext.Location = New-Object System.Drawing.Point(340, 12)
    $btnNext.Size = New-Object System.Drawing.Size(120, 34); $btnNext.BackColor = $Accent; $btnNext.ForeColor = $BgDark
    $btnNext.FlatStyle = "Flat"; $btnNext.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnNext.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $pnlBottom.Controls.Add($btnNext)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancelar"; $btnCancel.Location = New-Object System.Drawing.Point(480, 12)
    $btnCancel.Size = New-Object System.Drawing.Size(100, 34); $btnCancel.BackColor = $BgField; $btnCancel.ForeColor = $Red
    $btnCancel.FlatStyle = "Flat"; $btnCancel.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnCancel.Add_Click({ $wiz.Close() })
    $pnlBottom.Controls.Add($btnCancel)

    $wiz.Controls.Add($pnlBottom)

    # ── Logica de navegacao ──
    $script:wizStep = 0

    function Set-WizStep {
        param([int]$step)
        $script:wizStep = $step
        for ($i = 0; $i -lt $panels.Count; $i++) {
            $panels[$i].Visible = ($i -eq $step)
            $panels[$i].BringToFront()
        }
        # Highlight step labels
        for ($i = 0; $i -lt $stepLabels.Count; $i++) {
            if ($i -eq $step) {
                $stepLabels[$i].ForeColor = [System.Drawing.Color]::FromArgb(203,166,247)
                $stepLabels[$i].Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
            } elseif ($i -lt $step) {
                $stepLabels[$i].ForeColor = $Green
                $stepLabels[$i].Font = New-Object System.Drawing.Font("Segoe UI", 11)
            } else {
                $stepLabels[$i].ForeColor = [System.Drawing.Color]::FromArgb(108,112,134)
                $stepLabels[$i].Font = New-Object System.Drawing.Font("Segoe UI", 11)
            }
        }
        $btnPrev.Enabled = ($step -gt 0)
        if ($step -eq 3) { $btnNext.Text = "CRIAR GPO!"; $btnNext.BackColor = $Green }
        else { $btnNext.Text = "Proximo >"; $btnNext.BackColor = $Accent }

        # Montar resumo no step 4
        if ($step -eq 3) {
            $resumo = "NOME: $($txtWizName.Text.Trim())`n"
            $resumo += "DESCRICAO: $(if ($txtWizDesc.Text.Trim()) { $txtWizDesc.Text.Trim() } else { '(nenhuma)' })`n`n"
            $resumo += "POLITICAS SELECIONADAS:`n"
            $anyPol = $false
            foreach ($key in $script:wizChecks.Keys) {
                if ($script:wizChecks[$key].Checked) {
                    $resumo += "  [x] $($script:wizChecks[$key].Text)`n"
                    $anyPol = $true
                }
            }
            if (-not $anyPol) { $resumo += "  (nenhuma - voce pode adicionar depois)`n" }
            $resumo += "`nDESTINO (OUs):`n"
            $anyOU = $false
            $checkedOUs = Get-CheckedWizOUs $wizTree.Nodes
            foreach ($ou in $checkedOUs) {
                $resumo += "  [x] $ou`n"
                $anyOU = $true
            }
            if (-not $anyOU) { $resumo += "  (nenhum - voce pode vincular depois)`n" }
            $lblResume.Text = $resumo
        }
    }

    Set-WizStep 0

    $btnNext.Add_Click({
        if ($script:wizStep -eq 0) {
            # Validar nome
            if ([string]::IsNullOrEmpty($txtWizName.Text.Trim())) {
                Show-DarkMsg "Digite um nome para a GPO!" "Aviso" "OK" "Warning"
                return
            }
            Set-WizStep 1
        } elseif ($script:wizStep -eq 1) {
            Set-WizStep 2
        } elseif ($script:wizStep -eq 2) {
            Set-WizStep 3
        } elseif ($script:wizStep -eq 3) {
            # ═══ CRIAR GPO! ═══
            $wiz.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            $btnNext.Enabled = $false
            try {
                # 1) Criar GPO via ADSI
                $policiesPath = "LDAP://CN=Policies,CN=System,$domDN"
                $policies = [ADSI]$policiesPath
                $newGuid = [Guid]::NewGuid().ToString('B').ToUpper()
                $newGPO = $policies.Create("groupPolicyContainer", "CN=$newGuid")
                $newGPO.Put("displayName", $txtWizName.Text.Trim())
                $newGPO.Put("gPCFunctionalityVersion", 2)
                $newGPO.Put("flags", 0)
                $newGPO.Put("versionNumber", 0)
                $gpcPath = "\\$($script:Domain)\SysVol\$($script:Domain)\Policies\$newGuid"
                $newGPO.Put("gPCFileSysPath", $gpcPath)
                $newGPO.SetInfo()

                # 2) Criar pasta SYSVOL
                New-Item -Path $gpcPath -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
                New-Item -Path "$gpcPath\Machine" -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
                New-Item -Path "$gpcPath\User" -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null

                # 3) Salvar descricao
                if ($txtWizDesc.Text.Trim()) {
                    try {
                        $gpoEntry = [ADSI]"LDAP://CN=$newGuid,CN=Policies,CN=System,$domDN"
                        $gpoEntry.Put("description", $txtWizDesc.Text.Trim())
                        $gpoEntry.SetInfo()
                    } catch {}
                }

                # 4) Salvar politicas marcadas no SYSVOL
                # Mapeamento do wizard para IDs do editor
                $wizToEditor = @{
                    "wizBlkUSB"      = "DisableUSB"
                    "wizBlkCmd"      = "DisableCMD"
                    "wizBlkRegedit"  = "DisableRegistry"
                    "wizBlkPanel"    = "DisableControlPanel"
                    "wizBlkTaskMgr"  = "DisableTaskMgr"
                    "wizPwdPolicy"   = "MinPwdAge"
                    "wizNoTheme"     = "NoThemes"
                    "wizWallpaper"   = "NoWallpaper"
                    "wizNoUpdate"    = "DisableWindowsUpdate"
                    "wizNoStore"     = "DisableStore"
                    "wizNoCortana"   = "DisableCortana"
                    "wizNoOneDrive"  = "DisableOneDrive"
                    "wizNoAutoPlay"  = "NoAutoPlay"
                    "wizNoTelemetry" = "DisableTelemetry"
                    "wizMapDrive"    = "MapNetworkDrive"
                    "wizFirewall"    = "EnableFirewall"
                    "wizNeverSleep"  = "NoSleep"
                    "wizNoDesktop"   = "NoDispCPL"
                }
                $polConfig = @{}
                foreach ($key in $script:wizChecks.Keys) {
                    if ($script:wizChecks[$key].Checked) {
                        $editorId = if ($wizToEditor.ContainsKey($key)) { $wizToEditor[$key] } else { $key }
                        $polConfig[$editorId] = $true
                    }
                }
                if ($polConfig.Count -gt 0) {
                    $polConfig | ConvertTo-Json | Set-Content "$gpcPath\Machine\gpo_config.json" -Force
                }

                # 5) Vincular nas OUs marcadas
                $checkedOUs = Get-CheckedWizOUs $wizTree.Nodes
                $linkErrors = @()
                foreach ($ouDN in $checkedOUs) {
                    try {
                        $ouObj = [ADSI]"LDAP://$ouDN"
                        $curLink = [string]$ouObj.gPLink
                        $newLinkEntry = "[LDAP://CN=$newGuid,CN=Policies,CN=System,$domDN;0]"
                        if ($curLink) { $ouObj.Put("gPLink", "$curLink$newLinkEntry") }
                        else { $ouObj.Put("gPLink", $newLinkEntry) }
                        $ouObj.SetInfo()
                    } catch {
                        $linkErrors += "$ouDN : $($_.Exception.Message)"
                    }
                }

                $msg = "GPO '$($txtWizName.Text.Trim())' criada com sucesso!`n`n"
                $msg += "Politicas: $($polConfig.Count) selecionadas`n"
                $msg += "Destinos: $($checkedOUs.Count) OUs vinculadas`n"
                if ($linkErrors.Count -gt 0) { $msg += "`nErros de vinculo:`n" + ($linkErrors -join "`n") }
                $msg += "`n`nPara configurar detalhes, edite a GPO na tela principal."

                Show-DarkMsg $msg "GPO Criada!" "OK" "Information"
                Load-GPOs
                $wiz.Close()

            } catch {
                Show-DarkMsg "Erro ao criar GPO: $($_.Exception.Message)" "Falha" "OK" "Error"
            } finally {
                $wiz.Cursor = [System.Windows.Forms.Cursors]::Default
                $btnNext.Enabled = $true
            }
        }
    })

    $btnPrev.Add_Click({
        if ($script:wizStep -gt 0) { Set-WizStep ($script:wizStep - 1) }
    })

    $wiz.ShowDialog() | Out-Null
    $wiz.Dispose()
}

function Delete-SelectedGPO {
    if ($dgv.SelectedRows.Count -eq 0) {
        Show-DarkMsg "Selecione uma GPO." "Aviso" "OK" "Warning"; return
    }
    $name = $dgv.SelectedRows[0].Cells["Name"].Value
    $id   = $dgv.SelectedRows[0].Cells["Id"].Value

    $confirm = Show-DarkMsg "Excluir a GPO '$name'?`n`nEssa acao nao pode ser desfeita!" "Confirmar Exclusao" "YesNo" "Warning"

    if ($confirm -eq "Yes") {
        try {
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            # Excluir GPO via ADSI
            $domDN = ($script:Domain.Split('.') | ForEach-Object { "DC=$_" }) -join ','
            $gpoPath = "LDAP://CN=$id,CN=Policies,CN=System,$domDN"
            $gpoEntry = [ADSI]$gpoPath
            $sysvolPath = $gpoEntry.gPCFileSysPath
            $gpoEntry.DeleteTree()
            
            # Remover pasta SYSVOL
            if ($sysvolPath -and (Test-Path $sysvolPath)) {
                Remove-Item -Path $sysvolPath -Recurse -Force -ErrorAction SilentlyContinue
            }
            
            Show-DarkMsg "GPO '$name' excluida." "Sucesso" "OK" "Information"
            Load-GPOs
        } catch {
            Show-DarkMsg "Erro: $($_.Exception.Message)" "Falha" "OK" "Error"
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }
}

function Edit-SelectedGPO {
    if ($dgv.SelectedRows.Count -eq 0) {
        Show-DarkMsg "Selecione uma GPO para editar." "Aviso" "OK" "Warning"; return
    }
    $gpoId   = $dgv.SelectedRows[0].Cells["Id"].Value
    $gpoName = $dgv.SelectedRows[0].Cells["Name"].Value
    Edit-GPO -GpoId $gpoId -GpoName $gpoName
}

# ══════════════════════════════════════════
#  EXPORTAR GPO (salva config em JSON local)
# ══════════════════════════════════════════
function Export-SelectedGPO {
    if ($dgv.SelectedRows.Count -eq 0) {
        Show-DarkMsg "Selecione uma GPO para exportar." "Aviso" "OK" "Warning"; return
    }
    $gpoId   = $dgv.SelectedRows[0].Cells["Id"].Value
    $gpoName = $dgv.SelectedRows[0].Cells["Name"].Value

    try {
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $domDN = ($script:Domain.Split('.') | ForEach-Object { "DC=$_" }) -join ','
        $gpoEntry = [ADSI]"LDAP://CN=$gpoId,CN=Policies,CN=System,$domDN"
        $sysvolPath = [string]$gpoEntry.gPCFileSysPath
        $machPath = "$sysvolPath\Machine"

        $export = @{
            ExportVersion = 1
            ExportDate    = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            Name          = $gpoName
            Description   = ""
            Flags         = [int]$gpoEntry.flags.Value
            Policies      = @{}
            BlockedApps   = @()
            RegistryRules = @()
        }

        try { if ($gpoEntry.Properties["description"].Count -gt 0) { $export.Description = [string]$gpoEntry.Properties["description"][0] } } catch {}

        # Ler config de politicas
        $configFile = "$machPath\gpo_config.json"
        if (Test-Path $configFile) {
            try { $export.Policies = Get-Content $configFile -Raw | ConvertFrom-Json } catch {}
        }

        # Ler apps bloqueados
        $blockFile = "$machPath\blocked_apps.txt"
        if (Test-Path $blockFile) {
            $export.BlockedApps = @(Get-Content $blockFile | Where-Object { $_.Trim() } | ForEach-Object { $_.Trim() })
        }

        # Ler registry rules
        $regFile = "$machPath\registry_rules.json"
        if (Test-Path $regFile) {
            try { $export.RegistryRules = @(Get-Content $regFile -Raw | ConvertFrom-Json) } catch {}
        }

        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Title = "Exportar GPO - $gpoName"
        $sfd.Filter = "Arquivo GPO (*.gpo.json)|*.gpo.json"
        $sfd.FileName = "$($gpoName -replace '[^\w\-\. ]','_').gpo.json"
        $sfd.InitialDirectory = [Environment]::GetFolderPath("Desktop")

        if ($sfd.ShowDialog() -eq "OK") {
            $export | ConvertTo-Json -Depth 10 | Set-Content $sfd.FileName -Encoding UTF8
            Show-DarkMsg "GPO '$gpoName' exportada com sucesso!`n`n$($sfd.FileName)" "Exportar" "OK" "Information"
        }
    } catch {
        Show-DarkMsg "Erro ao exportar: $($_.Exception.Message)" "Erro" "OK" "Error"
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
}

# ══════════════════════════════════════════
#  IMPORTAR GPO (cria nova a partir de JSON)
# ══════════════════════════════════════════
function Import-GPOFromFile {
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Title = "Importar GPO de arquivo"
    $ofd.Filter = "Arquivo GPO (*.gpo.json)|*.gpo.json|JSON (*.json)|*.json"
    $ofd.InitialDirectory = [Environment]::GetFolderPath("Desktop")

    if ($ofd.ShowDialog() -ne "OK") { return }

    try {
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $import = Get-Content $ofd.FileName -Raw -Encoding UTF8 | ConvertFrom-Json

        if (-not $import.Name) {
            Show-DarkMsg "Arquivo invalido: campo 'Name' ausente." "Erro" "OK" "Error"; return
        }

        $newName = "$($import.Name) (importada)"
        $conf = Show-DarkMsg "Importar GPO como:`n$newName`n`nPoliticas: $(if ($import.Policies) { 'Sim' } else { 'Nao' })`nApps bloqueados: $(@($import.BlockedApps).Count)`nRegras registro: $(@($import.RegistryRules).Count)`n`nContinuar?" "Importar GPO" "YesNo" "Question"
        if ($conf -ne "Yes") { return }

        # Criar GPO via ADSI
        $domDN = ($script:Domain.Split('.') | ForEach-Object { "DC=$_" }) -join ','
        $policiesPath = "LDAP://CN=Policies,CN=System,$domDN"
        $policies = [ADSI]$policiesPath
        $newGuid = [Guid]::NewGuid().ToString('B').ToUpper()
        $newGPO = $policies.Create("groupPolicyContainer", "CN=$newGuid")
        $newGPO.Put("displayName", $newName)
        $newGPO.Put("gPCFunctionalityVersion", 2)
        $flags = if ($import.Flags) { [int]$import.Flags } else { 0 }
        $newGPO.Put("flags", $flags)
        $newGPO.Put("versionNumber", 0)
        $gpcPath = "\\$($script:Domain)\SysVol\$($script:Domain)\Policies\$newGuid"
        $newGPO.Put("gPCFileSysPath", $gpcPath)
        if ($import.Description) { $newGPO.Put("description", $import.Description) }
        $newGPO.SetInfo()

        # Criar pasta SYSVOL e copiar configs
        $machPath = "$gpcPath\Machine"
        New-Item -Path $gpcPath -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
        New-Item -Path $machPath -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
        New-Item -Path "$gpcPath\User" -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null

        if ($import.Policies) {
            $import.Policies | ConvertTo-Json -Depth 5 | Out-File "$machPath\gpo_config.json" -Encoding UTF8
        }
        if ($import.BlockedApps -and @($import.BlockedApps).Count -gt 0) {
            $import.BlockedApps | Out-File "$machPath\blocked_apps.txt" -Encoding UTF8
        }
        if ($import.RegistryRules -and @($import.RegistryRules).Count -gt 0) {
            $import.RegistryRules | ConvertTo-Json -Depth 5 | Out-File "$machPath\registry_rules.json" -Encoding UTF8
        }

        Show-DarkMsg "GPO '$newName' importada com sucesso!" "Importar" "OK" "Information"
        Load-GPOs
    } catch {
        Show-DarkMsg "Erro ao importar: $($_.Exception.Message)" "Erro" "OK" "Error"
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
}

# ══════════════════════════════════════════
#  DUPLICAR GPO (clonar config)
# ══════════════════════════════════════════
function Clone-SelectedGPO {
    if ($dgv.SelectedRows.Count -eq 0) {
        Show-DarkMsg "Selecione uma GPO para duplicar." "Aviso" "OK" "Warning"; return
    }
    $gpoId   = $dgv.SelectedRows[0].Cells["Id"].Value
    $gpoName = $dgv.SelectedRows[0].Cells["Name"].Value

    $conf = Show-DarkMsg "Duplicar a GPO '$gpoName'?`n`nSera criada uma copia com todas as configuracoes." "Duplicar GPO" "YesNo" "Question"
    if ($conf -ne "Yes") { return }

    try {
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $domDN = ($script:Domain.Split('.') | ForEach-Object { "DC=$_" }) -join ','

        # Ler GPO original
        $srcEntry = [ADSI]"LDAP://CN=$gpoId,CN=Policies,CN=System,$domDN"
        $srcSysvol = [string]$srcEntry.gPCFileSysPath
        $srcFlags = [int]$srcEntry.flags.Value
        $srcDesc = ""
        try { if ($srcEntry.Properties["description"].Count -gt 0) { $srcDesc = [string]$srcEntry.Properties["description"][0] } } catch {}

        # Criar nova GPO
        $newName = "$gpoName (copia)"
        $policiesPath = "LDAP://CN=Policies,CN=System,$domDN"
        $policies = [ADSI]$policiesPath
        $newGuid = [Guid]::NewGuid().ToString('B').ToUpper()
        $newGPO = $policies.Create("groupPolicyContainer", "CN=$newGuid")
        $newGPO.Put("displayName", $newName)
        $newGPO.Put("gPCFunctionalityVersion", 2)
        $newGPO.Put("flags", $srcFlags)
        $newGPO.Put("versionNumber", 0)
        $gpcPath = "\\$($script:Domain)\SysVol\$($script:Domain)\Policies\$newGuid"
        $newGPO.Put("gPCFileSysPath", $gpcPath)
        if ($srcDesc) { $newGPO.Put("description", $srcDesc) }
        $newGPO.SetInfo()

        # Copiar pasta SYSVOL inteira
        New-Item -Path $gpcPath -ItemType Directory -Force | Out-Null
        if ($srcSysvol -and (Test-Path $srcSysvol)) {
            Copy-Item -Path "$srcSysvol\*" -Destination $gpcPath -Recurse -Force -ErrorAction SilentlyContinue
        } else {
            New-Item -Path "$gpcPath\Machine" -ItemType Directory -Force | Out-Null
            New-Item -Path "$gpcPath\User" -ItemType Directory -Force | Out-Null
        }

        Show-DarkMsg "GPO '$newName' criada como copia de '$gpoName'!" "Duplicar" "OK" "Information"
        Load-GPOs
    } catch {
        Show-DarkMsg "Erro ao duplicar: $($_.Exception.Message)" "Erro" "OK" "Error"
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
}

# Helper: buscar todos OUs do AD
function Get-AllOUs {
    param([string]$DomDN)
    $list = [System.Collections.ArrayList]@()
    try {
        $s = New-Object System.DirectoryServices.DirectorySearcher
        $s.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$DomDN")
        $s.Filter = "(objectClass=organizationalUnit)"; $s.PageSize = 1000
        $s.PropertiesToLoad.AddRange(@("distinguishedName","name","description")) | Out-Null
        $s.SearchScope = "Subtree"
        $res = $s.FindAll()
        foreach ($r in $res) {
            $list.Add([PSCustomObject]@{
                DN   = [string]$r.Properties["distinguishedname"][0]
                Name = [string]$r.Properties["name"][0]
                Desc = if ($r.Properties["description"].Count -gt 0) { [string]$r.Properties["description"][0] } else { "" }
            }) | Out-Null
        }
        $res.Dispose(); $s.Dispose()
    } catch {}
    return $list
}

# Helper: buscar todos Grupos do AD
function Get-AllADGroups {
    param([string]$DomDN)
    $list = [System.Collections.ArrayList]@()
    try {
        $s = New-Object System.DirectoryServices.DirectorySearcher
        $s.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$DomDN")
        $s.Filter = "(objectClass=group)"; $s.PageSize = 1000
        $s.PropertiesToLoad.AddRange(@("distinguishedName","cn","description","groupType","member")) | Out-Null
        $res = $s.FindAll()
        foreach ($r in $res) {
            $gt = if ($r.Properties["grouptype"].Count -gt 0) { [int]$r.Properties["grouptype"][0] } else { 0 }
            $tipo = if ($gt -band 0x80000000) { "Seguranca" } else { "Distribuicao" }
            $memberCount = $r.Properties["member"].Count
            $list.Add([PSCustomObject]@{
                DN      = [string]$r.Properties["distinguishedname"][0]
                Name    = [string]$r.Properties["cn"][0]
                Desc    = if ($r.Properties["description"].Count -gt 0) { [string]$r.Properties["description"][0] } else { "" }
                Type    = $tipo
                Members = $memberCount
            }) | Out-Null
        }
        $res.Dispose(); $s.Dispose()
    } catch {}
    return $list
}

# Helper: buscar computadores do AD
function Get-AllADComputers {
    param([string]$DomDN)
    $list = [System.Collections.ArrayList]@()
    try {
        $s = New-Object System.DirectoryServices.DirectorySearcher
        $s.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$DomDN")
        $s.Filter = "(objectClass=computer)"; $s.PageSize = 1000
        $s.PropertiesToLoad.AddRange(@("distinguishedName","cn","operatingSystem","lastLogonTimestamp","description")) | Out-Null
        $res = $s.FindAll()
        foreach ($r in $res) {
            $os = if ($r.Properties["operatingsystem"].Count -gt 0) { [string]$r.Properties["operatingsystem"][0] } else { "" }
            $lastLogon = ""
            if ($r.Properties["lastlogontimestamp"].Count -gt 0) {
                try { $lastLogon = [datetime]::FromFileTime([long]$r.Properties["lastlogontimestamp"][0]).ToString("yyyy-MM-dd HH:mm") } catch {}
            }
            $list.Add([PSCustomObject]@{
                DN       = [string]$r.Properties["distinguishedname"][0]
                Name     = [string]$r.Properties["cn"][0]
                OS       = $os
                LastLogon= $lastLogon
                Desc     = if ($r.Properties["description"].Count -gt 0) { [string]$r.Properties["description"][0] } else { "" }
            }) | Out-Null
        }
        $res.Dispose(); $s.Dispose()
    } catch {}
    return $list
}

# Helper: buscar usuarios do AD
function Get-AllADUsers {
    param([string]$DomDN)
    $list = [System.Collections.ArrayList]@()
    try {
        $s = New-Object System.DirectoryServices.DirectorySearcher
        $s.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$DomDN")
        $s.Filter = "(&(objectClass=user)(objectCategory=person))"; $s.PageSize = 1000
        $s.PropertiesToLoad.AddRange(@("distinguishedName","cn","sAMAccountName","mail","department","title","lastLogonTimestamp","userAccountControl")) | Out-Null
        $res = $s.FindAll()
        foreach ($r in $res) {
            $uac = if ($r.Properties["useraccountcontrol"].Count -gt 0) { [int]$r.Properties["useraccountcontrol"][0] } else { 0 }
            $enabled = -not ($uac -band 2)
            $lastLogon = ""
            if ($r.Properties["lastlogontimestamp"].Count -gt 0) {
                try { $lastLogon = [datetime]::FromFileTime([long]$r.Properties["lastlogontimestamp"][0]).ToString("yyyy-MM-dd HH:mm") } catch {}
            }
            $list.Add([PSCustomObject]@{
                DN       = [string]$r.Properties["distinguishedname"][0]
                Name     = [string]$r.Properties["cn"][0]
                Login    = if ($r.Properties["samaccountname"].Count -gt 0) { [string]$r.Properties["samaccountname"][0] } else { "" }
                Email    = if ($r.Properties["mail"].Count -gt 0) { [string]$r.Properties["mail"][0] } else { "" }
                Dept     = if ($r.Properties["department"].Count -gt 0) { [string]$r.Properties["department"][0] } else { "" }
                Title    = if ($r.Properties["title"].Count -gt 0) { [string]$r.Properties["title"][0] } else { "" }
                Enabled  = $enabled
                LastLogon= $lastLogon
            }) | Out-Null
        }
        $res.Dispose(); $s.Dispose()
    } catch {}
    return $list
}

function Edit-GPO {
    param([string]$GpoId, [string]$GpoName)

    try { Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction SilentlyContinue } catch {}

    try {
        $domDN = ($script:Domain.Split('.') | ForEach-Object { "DC=$_" }) -join ','
        $gpoEntry = [ADSI]"LDAP://CN=$GpoId,CN=Policies,CN=System,$domDN"
        $sysvolPath = [string]$gpoEntry.gPCFileSysPath
    } catch {
        Show-DarkMsg "Erro ao acessar GPO: $($_.Exception.Message)" "Erro" "OK" "Error"
        return
    }

    # ── Ler config salva no SYSVOL ──
    $script:savedPolicies = @{}
    $configFile = "$sysvolPath\Machine\gpo_config.json"
    if ($sysvolPath -and (Test-Path $configFile)) {
        try {
            $jsonObj = Get-Content $configFile -Raw | ConvertFrom-Json -ErrorAction Stop
            if ($jsonObj.PSObject) {
                foreach ($p in $jsonObj.PSObject.Properties) {
                    $script:savedPolicies[$p.Name] = $p.Value
                }
            }
        } catch {}
    }

    # ── Ler Registry.pol nativo (configuracoes GPMC) ──
    $script:nativeRegEntries = @()
    function Parse-RegistryPol {
        param([string]$PolPath, [string]$Hive)
        $entries = @()
        try {
            $bytes = [System.IO.File]::ReadAllBytes($PolPath)
            if ($bytes.Length -lt 8) { return $entries }
            $sig = [System.Text.Encoding]::ASCII.GetString($bytes, 0, 4)
            if ($sig -ne "PReg") { return $entries }
            $pos = 8
            while ($pos -lt ($bytes.Length - 10)) {
                if ($bytes[$pos] -ne 0x5B -or $bytes[$pos+1] -ne 0x00) { $pos++; continue }
                $pos += 2
                $start = $pos
                while ($pos -lt ($bytes.Length - 1) -and -not ($bytes[$pos] -eq 0x3B -and $bytes[$pos+1] -eq 0x00)) { $pos += 2 }
                $key = [System.Text.Encoding]::Unicode.GetString($bytes, $start, $pos - $start).TrimEnd([char]0)
                $pos += 2
                $start = $pos
                while ($pos -lt ($bytes.Length - 1) -and -not ($bytes[$pos] -eq 0x3B -and $bytes[$pos+1] -eq 0x00)) { $pos += 2 }
                $valName = [System.Text.Encoding]::Unicode.GetString($bytes, $start, $pos - $start).TrimEnd([char]0)
                $pos += 2
                if (($pos + 4) -gt $bytes.Length) { break }
                $type = [BitConverter]::ToUInt32($bytes, $pos); $pos += 4
                $pos += 2
                if (($pos + 4) -gt $bytes.Length) { break }
                $size = [BitConverter]::ToUInt32($bytes, $pos); $pos += 4
                $pos += 2
                if ($size -gt 0 -and ($pos + $size) -le $bytes.Length) {
                    $dataBytes = $bytes[$pos..($pos+$size-1)]
                } else { $dataBytes = @() }
                $pos += $size
                $pos += 2
                $typeName = switch ($type) { 1 {"String"} 2 {"ExpandString"} 4 {"DWord"} 7 {"MultiString"} 11 {"QWord"} 3 {"Binary"} default {"Tipo$type"} }
                $dataStr = switch ($type) {
                    4 { if($dataBytes.Count -ge 4){[string][BitConverter]::ToUInt32($dataBytes, 0)}else{"?"} }
                    11 { if($dataBytes.Count -ge 8){[string][BitConverter]::ToUInt64($dataBytes, 0)}else{"?"} }
                    3 { "(" + $dataBytes.Count + " bytes)" }
                    default { [System.Text.Encoding]::Unicode.GetString($dataBytes).TrimEnd([char]0) }
                }
                if ($key -and $valName -and $type -ne 0) {
                    $entries += @{ Hive=$Hive; Key=$key; ValueName=$valName; ValueData=$dataStr; Type=$typeName; Source="Registry.pol" }
                }
            }
        } catch {}
        return $entries
    }

    # Machine\Registry.pol
    foreach ($machSub in @("Machine","MACHINE")) {
        $polFile = "$sysvolPath\$machSub\Registry.pol"
        if (Test-Path $polFile) {
            $script:nativeRegEntries += Parse-RegistryPol $polFile "HKLM"
            break
        }
    }
    # User\Registry.pol
    foreach ($userSub in @("User","USER")) {
        $polFile = "$sysvolPath\$userSub\Registry.pol"
        if (Test-Path $polFile) {
            $script:nativeRegEntries += Parse-RegistryPol $polFile "HKCU"
            break
        }
    }

    # ── Ler scripts.ini nativo ──
    $script:nativeScripts = @()
    foreach ($sub in @("Machine","MACHINE")) {
        $iniPath = "$sysvolPath\$sub\Scripts\scripts.ini"
        if (Test-Path $iniPath) {
            try {
                $iniContent = Get-Content $iniPath -Encoding Unicode -ErrorAction SilentlyContinue
                $section = ""
                $idx = 0
                foreach ($line in $iniContent) {
                    if ($line -match '^\[(.+)\]') { $section = $Matches[1]; $idx = 0; continue }
                    if ($line -match "^${idx}CmdLine=(.+)") {
                        $cmd = $Matches[1].Trim()
                        $parLine = $iniContent | Where-Object { $_ -match "^${idx}Parameters=" } | Select-Object -First 1
                        $par = if ($parLine -match "=(.*)") { $Matches[1].Trim() } else { "" }
                        $script:nativeScripts += @{ Type=$section; Command=$cmd; Parameters=$par }
                        $idx++
                    }
                }
            } catch {}
            break
        }
    }

    # ── Ler GPT.INI ──
    $script:gptVersion = ""
    $gptIni = "$sysvolPath\GPT.INI"
    if (Test-Path $gptIni) {
        try {
            $content = Get-Content $gptIni -ErrorAction SilentlyContinue
            $vLine = $content | Where-Object { $_ -match "^Version=" } | Select-Object -First 1
            if ($vLine -match "=(\d+)") { $script:gptVersion = $Matches[1] }
        } catch {}
    }

    # ── Ler Preferences XML (GP Preferences nativas) ──
    $script:nativePreferences = @()
    $actionMap = @{ "C"="Criar"; "U"="Atualizar"; "R"="Substituir"; "D"="Deletar" }
    foreach ($scope in @("Machine","User","MACHINE","USER")) {
        $prefRoot = "$sysvolPath\$scope\Preferences"
        if (-not (Test-Path $prefRoot)) { continue }
        $scopeLabel = if ($scope -match "^[Mm]") { "Computador" } else { "Usuario" }
        Get-ChildItem $prefRoot -Directory -ErrorAction SilentlyContinue | ForEach-Object {
            $prefType = $_.Name
            Get-ChildItem $_.FullName -Filter "*.xml" -ErrorAction SilentlyContinue | ForEach-Object {
                try {
                    $xml = [xml](Get-Content $_.FullName -Raw -ErrorAction SilentlyContinue)
                    if (-not $xml) { return }
                    $root = $xml.DocumentElement
                    foreach ($node in $root.ChildNodes) {
                        if ($node.NodeType -ne "Element") { continue }
                        $act = if ($node.Properties) { $node.Properties.action } else { "" }
                        $actLabel = if ($actionMap[$act]) { $actionMap[$act] } else { $act }
                        switch ($prefType) {
                            "Shortcuts" {
                                $p = $node.Properties
                                $script:nativePreferences += @{
                                    Tipo="Atalho"; Escopo=$scopeLabel; Acao=$actLabel
                                    Nome=$node.name; Caminho=$p.targetPath
                                    Detalhes="Em: $($p.shortcutPath)  Icone: $($p.iconPath)  Args: $($p.arguments)"
                                }
                            }
                            "Drives" {
                                $p = $node.Properties
                                $script:nativePreferences += @{
                                    Tipo="Drive Mapeado"; Escopo=$scopeLabel; Acao=$actLabel
                                    Nome="$($p.letter): $($node.name)"; Caminho=$p.path
                                    Detalhes="Label: $($p.label)  Persistente: $($p.persistent)"
                                }
                            }
                            "Printers" {
                                $p = $node.Properties
                                $pName = if ($p.localName) { $p.localName } else { $node.name }
                                $pPath = if ($p.path) { $p.path } else { $p.ipAddress }
                                $script:nativePreferences += @{
                                    Tipo="Impressora"; Escopo=$scopeLabel; Acao=$actLabel
                                    Nome=$pName; Caminho=$pPath
                                    Detalhes="Padrao: $($p.default)  IP: $($p.ipAddress)"
                                }
                            }
                            "Files" {
                                $p = $node.Properties
                                $script:nativePreferences += @{
                                    Tipo="Arquivo"; Escopo=$scopeLabel; Acao=$actLabel
                                    Nome=$node.name; Caminho="$($p.fromPath) -> $($p.targetPath)"
                                    Detalhes="Oculto: $($p.hidden)  SomenteLeitura: $($p.readOnly)"
                                }
                            }
                            "Folders" {
                                $p = $node.Properties
                                $script:nativePreferences += @{
                                    Tipo="Pasta"; Escopo=$scopeLabel; Acao=$actLabel
                                    Nome=$node.name; Caminho=$p.path
                                    Detalhes="Oculto: $($p.hidden)  SomenteLeitura: $($p.readOnly)"
                                }
                            }
                            "Registry" {
                                $p = $node.Properties
                                $script:nativePreferences += @{
                                    Tipo="Registro (Pref)"; Escopo=$scopeLabel; Acao=$actLabel
                                    Nome="$($p.hive)\$($p.key)\$($p.name)"; Caminho="$($p.type) = $($p.value)"
                                    Detalhes="Hive: $($p.hive)"
                                }
                            }
                            "Services" {
                                $p = $node.Properties
                                $script:nativePreferences += @{
                                    Tipo="Servico"; Escopo=$scopeLabel; Acao=$actLabel
                                    Nome=$p.serviceName; Caminho="Inicio: $($p.startupType)"
                                    Detalhes="Acao: $($p.serviceAction)  Nome: $($node.name)"
                                }
                            }
                            "Groups" {
                                $p = $node.Properties
                                $gName = if ($p.groupName) { $p.groupName } else { $node.name }
                                $script:nativePreferences += @{
                                    Tipo="Grupo/Usuario"; Escopo=$scopeLabel; Acao=$actLabel
                                    Nome=$gName; Caminho=""
                                    Detalhes="Nome: $($node.name)"
                                }
                            }
                            "PowerOptions" {
                                $script:nativePreferences += @{
                                    Tipo="Energia"; Escopo=$scopeLabel; Acao=$actLabel
                                    Nome=$node.name; Caminho=""
                                    Detalhes=""
                                }
                            }
                            default {
                                $script:nativePreferences += @{
                                    Tipo=$prefType; Escopo=$scopeLabel; Acao=$actLabel
                                    Nome=$node.name; Caminho=""
                                    Detalhes=""
                                }
                            }
                        }
                    }
                } catch {}
            }
        }
    }

    # ── Ler fdeploy.ini (Redirecionamento de Pastas) ──
    $script:nativeFolderRedirect = @()
    $folderStatusMap = @{ "0"="Basico (redirecionar para local)"; "10"="Basico (redirecionar de volta)"; "4"="Avancado" }
    foreach ($scope in @("User","USER")) {
        $fdeployPath = "$sysvolPath\$scope\Documents & Settings\fdeploy.ini"
        if (Test-Path $fdeployPath) {
            try {
                $fContent = Get-Content $fdeployPath -Encoding Unicode -ErrorAction SilentlyContinue
                $section = ""; $statusMap = @{}
                foreach ($line in $fContent) {
                    $tl = $line.Trim()
                    if ($tl -match '^\[(.+)\]$') { $section = $Matches[1]; continue }
                    if (-not $tl -or $tl.StartsWith(";")) { continue }
                    if ($section -eq "FolderStatus" -and $tl -match '^(.+?)=(\d+)') {
                        $statusMap[$Matches[1]] = $Matches[2]
                    }
                    elseif ($section -ne "FolderStatus" -and $tl -match '^(.+?)=(.+)$') {
                        $sid = $Matches[1]; $target = $Matches[2]
                        $sidLabel = switch -Wildcard ($sid) { "s-1-1-0" {"Todos"}; "S-1-1-0" {"Todos"}; default {$sid} }
                        $stVal = if ($statusMap[$section]) { $statusMap[$section] } else { "?" }
                        $stLabel = if ($folderStatusMap[$stVal]) { $folderStatusMap[$stVal] } else { "Status=$stVal" }
                        $script:nativeFolderRedirect += @{
                            Pasta=$section; Destino=$target; Aplica=$sidLabel; Status=$stLabel
                        }
                        $script:nativePreferences += @{
                            Tipo="Redir. Pasta"; Escopo="Usuario"; Acao=$stLabel
                            Nome=$section; Caminho=$target; Detalhes="Aplica: $sidLabel"
                        }
                    }
                }
            } catch {}
            break
        }
    }

    # ── Ler GptTmpl.inf (Politicas de Seguranca) ──
    $script:nativeSecuritySettings = @()
    foreach ($sub in @("Machine","MACHINE")) {
        $tmplPath = "$sysvolPath\$sub\Microsoft\Windows NT\SecEdit\GptTmpl.inf"
        if (Test-Path $tmplPath) {
            try {
                $tmplContent = Get-Content $tmplPath -ErrorAction SilentlyContinue
                $section = ""
                foreach ($line in $tmplContent) {
                    $tl = $line.Trim()
                    if ($tl -match '^\[(.+)\]$') { $section = $Matches[1]; continue }
                    if (-not $tl -or $tl.StartsWith(";") -or $section -eq "Unicode" -or $section -eq "Version") { continue }
                    switch ($section) {
                        "System Access" {
                            if ($tl -match '^(.+?)\s*=\s*(.+)$') {
                                $settingNames = @{
                                    "MinimumPasswordAge"="Idade minima da senha (dias)"
                                    "MaximumPasswordAge"="Idade maxima da senha (dias)"
                                    "MinimumPasswordLength"="Comprimento minimo da senha"
                                    "PasswordHistorySize"="Historico de senhas"
                                    "LockoutBadCount"="Tentativas antes do bloqueio"
                                    "RequireLogonToChangePassword"="Exigir logon para alterar senha"
                                    "ForceLogoffWhenHourExpire"="Forcar logoff ao expirar horario"
                                    "ClearTextPassword"="Armazenar senhas em texto limpo"
                                    "LSAAnonymousNameLookup"="Permitir consulta anonima LSA"
                                    "PasswordComplexity"="Exigir complexidade de senha"
                                    "LockoutDuration"="Duracao do bloqueio (min)"
                                    "ResetLockoutCount"="Resetar contador apos (min)"
                                }
                                $name = if ($settingNames[$Matches[1]]) { $settingNames[$Matches[1]] } else { $Matches[1] }
                                $entry = @{ Secao="Acesso ao Sistema"; Nome=$name; Valor=$Matches[2] }
                                $script:nativeSecuritySettings += $entry
                                $script:nativePreferences += @{
                                    Tipo="Seguranca"; Escopo="Computador"; Acao="Configurado"
                                    Nome=$name; Caminho=$Matches[2]; Detalhes="[System Access] $($Matches[1])"
                                }
                            }
                        }
                        "Kerberos Policy" {
                            if ($tl -match '^(.+?)\s*=\s*(.+)$') {
                                $kNames = @{
                                    "MaxTicketAge"="Validade maxima do ticket (hrs)"
                                    "MaxRenewAge"="Renovacao maxima do ticket (dias)"
                                    "MaxServiceAge"="Validade maxima do ticket de servico (min)"
                                    "MaxClockSkew"="Tolerancia de relogio (min)"
                                    "TicketValidateClient"="Validar cliente"
                                }
                                $name = if ($kNames[$Matches[1]]) { $kNames[$Matches[1]] } else { $Matches[1] }
                                $script:nativeSecuritySettings += @{ Secao="Kerberos"; Nome=$name; Valor=$Matches[2] }
                                $script:nativePreferences += @{
                                    Tipo="Seguranca"; Escopo="Computador"; Acao="Configurado"
                                    Nome=$name; Caminho=$Matches[2]; Detalhes="[Kerberos Policy]"
                                }
                            }
                        }
                        "Privilege Rights" {
                            if ($tl -match '^(.+?)\s*=\s*(.+)$') {
                                $privNames = @{
                                    "SeAssignPrimaryTokenPrivilege"="Substituir token de processo"
                                    "SeAuditPrivilege"="Gerar auditorias de seguranca"
                                    "SeBackupPrivilege"="Fazer backup de arquivos"
                                    "SeBatchLogonRight"="Logon como tarefa em lotes"
                                    "SeChangeNotifyPrivilege"="Ignorar verificacao transversal"
                                    "SeDebugPrivilege"="Depurar programas"
                                    "SeInteractiveLogonRight"="Logon local"
                                    "SeLoadDriverPrivilege"="Carregar drivers de dispositivo"
                                    "SeMachineAccountPrivilege"="Adicionar estacoes ao dominio"
                                    "SeNetworkLogonRight"="Acessar pela rede"
                                    "SeRemoteShutdownPrivilege"="Forcar desligamento remoto"
                                    "SeRestorePrivilege"="Restaurar arquivos"
                                    "SeSecurityPrivilege"="Gerenciar log de auditoria"
                                    "SeShutdownPrivilege"="Desligar o sistema"
                                    "SeSystemEnvironmentPrivilege"="Modificar valores de firmware"
                                    "SeSystemTimePrivilege"="Alterar hora do sistema"
                                    "SeTakeOwnershipPrivilege"="Apropriar-se de objetos"
                                    "SeUndockPrivilege"="Remover estacao de ancoragem"
                                    "SeEnableDelegationPrivilege"="Habilitar delegacao"
                                    "SeCreatePagefilePrivilege"="Criar arquivo de paginacao"
                                    "SeIncreaseBasePriorityPrivilege"="Aumentar prioridade"
                                    "SeIncreaseQuotaPrivilege"="Ajustar cotas de memoria"
                                    "SeProfileSingleProcessPrivilege"="Tracar perfil de processo"
                                    "SeSystemProfilePrivilege"="Tracar perfil do sistema"
                                }
                                $name = if ($privNames[$Matches[1]]) { $privNames[$Matches[1]] } else { $Matches[1] }
                                $sids = ($Matches[2] -split ',') | ForEach-Object { $_.Trim().TrimStart('*') }
                                $count = $sids.Count
                                $script:nativeSecuritySettings += @{ Secao="Direitos"; Nome=$name; Valor="$count SID(s)" }
                                $script:nativePreferences += @{
                                    Tipo="Privilegio"; Escopo="Computador"; Acao="$count SID(s)"
                                    Nome=$name; Caminho=($sids | Select -First 3) -join ", "
                                    Detalhes=if($count -gt 3){"+ $($count-3) mais..."}else{""}
                                }
                            }
                        }
                        "Group Membership" {
                            if ($tl -match '^(.+?)__(.+?)\s*=\s*(.*)$') {
                                $grp = $Matches[1].TrimStart('*'); $rel = $Matches[2]; $members = $Matches[3]
                                $script:nativeSecuritySettings += @{ Secao="Membros de Grupo"; Nome="$grp ($rel)"; Valor=$members }
                                $script:nativePreferences += @{
                                    Tipo="Membro Grupo"; Escopo="Computador"; Acao=$rel
                                    Nome=$grp; Caminho=$members; Detalhes=""
                                }
                            }
                        }
                        "Registry Values" {
                            if ($tl -match '^(.+?)\s*=\s*(\d+),(.+)$') {
                                $script:nativeSecuritySettings += @{ Secao="Valores de Registro"; Nome=$Matches[1]; Valor="Tipo$($Matches[2])=$($Matches[3])" }
                                $script:nativePreferences += @{
                                    Tipo="Reg. Seguranca"; Escopo="Computador"; Acao="Configurado"
                                    Nome=$Matches[1].Replace("MACHINE\",""); Caminho=$Matches[3]; Detalhes="RegType=$($Matches[2])"
                                }
                            }
                        }
                        "Registry Keys" {
                            if ($tl -match '^"(.+?)"') {
                                $script:nativePreferences += @{
                                    Tipo="Perm. Registro"; Escopo="Computador"; Acao="ACL"
                                    Nome=$Matches[1].Replace("MACHINE\",""); Caminho=""; Detalhes="Permissoes de chave"
                                }
                            }
                        }
                        "File Security" {
                            if ($tl -match '^"(.+?)"') {
                                $script:nativePreferences += @{
                                    Tipo="Perm. Arquivo"; Escopo="Computador"; Acao="ACL"
                                    Nome=$Matches[1]; Caminho=""; Detalhes="Permissoes de arquivo"
                                }
                            }
                        }
                    }
                }
            } catch {}
            break
        }
    }

    # ── Ler Software Deployment (.aas) ──
    foreach ($sub in @("Machine","MACHINE")) {
        $appsDir = "$sysvolPath\$sub\Applications"
        if (Test-Path $appsDir) {
            Get-ChildItem $appsDir -Filter "*.aas" -ErrorAction SilentlyContinue | ForEach-Object {
                $script:nativePreferences += @{
                    Tipo="Software (MSI)"; Escopo="Computador"; Acao="Implantado"
                    Nome=$_.BaseName; Caminho=$_.FullName; Detalhes="$($_.Length) bytes"
                }
            }
            break
        }
    }

    # ── Mapear Registry.pol sera feito apos $commonPolicies ──

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Editar GPO: $GpoName"
    $dlg.Size = New-Object System.Drawing.Size(1000, 750)
    $dlg.StartPosition = "CenterScreen"
    $dlg.BackColor = $BgDark
    $dlg.ForeColor = $FgText
    $dlg.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $dlg.MinimumSize = New-Object System.Drawing.Size(950, 700)
    $dlg.WindowState = "Maximized"

    $tabs = New-Object System.Windows.Forms.TabControl
    $tabs.Dock = "Fill"
    $tabs.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)

    # ════════════════════════════════════════
    #  TAB 1: GERAL - TODAS AS INFORMACOES
    # ════════════════════════════════════════
    $tabGeral = New-Object System.Windows.Forms.TabPage
    $tabGeral.Text = "  Geral  "
    $tabGeral.BackColor = $BgDark

    $pnlGeral = New-Object System.Windows.Forms.Panel
    $pnlGeral.Dock = "Fill"; $pnlGeral.AutoScroll = $true; $pnlGeral.Padding = New-Object System.Windows.Forms.Padding(20)

    $yy = 15

    # SECAO: Identificacao
    $sh = New-Object System.Windows.Forms.Label; $sh.Text = "IDENTIFICACAO"; $sh.AutoSize = $true
    $sh.ForeColor = $Accent; $sh.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $sh.Location = New-Object System.Drawing.Point(15, $yy); $pnlGeral.Controls.Add($sh); $yy += 30

    # Nome
    $lName = New-Object System.Windows.Forms.Label; $lName.Text = "Nome da GPO:"; $lName.AutoSize = $true
    $lName.Location = New-Object System.Drawing.Point(15, $yy); $pnlGeral.Controls.Add($lName); $yy += 22
    $tName = New-Object System.Windows.Forms.TextBox; $tName.Text = $GpoName
    $tName.Location = New-Object System.Drawing.Point(15, $yy); $tName.Size = New-Object System.Drawing.Size(900, 30)
    $tName.BackColor = $BgField; $tName.ForeColor = $FgText; $tName.Font = New-Object System.Drawing.Font("Segoe UI", 12)
    $tName.BorderStyle = "FixedSingle"
    $pnlGeral.Controls.Add($tName); $yy += 40

    # Descricao
    $lDesc = New-Object System.Windows.Forms.Label; $lDesc.Text = "Descricao / Comentario:"; $lDesc.AutoSize = $true
    $lDesc.Location = New-Object System.Drawing.Point(15, $yy); $pnlGeral.Controls.Add($lDesc); $yy += 22
    $curDesc = ""
    try { if ($gpoEntry.Properties["description"].Count -gt 0) { $curDesc = [string]$gpoEntry.Properties["description"][0] } } catch {}
    $tDesc = New-Object System.Windows.Forms.TextBox; $tDesc.Text = $curDesc; $tDesc.Multiline = $true
    $tDesc.Location = New-Object System.Drawing.Point(15, $yy); $tDesc.Size = New-Object System.Drawing.Size(900, 50)
    $tDesc.BackColor = $BgField; $tDesc.ForeColor = $FgText; $tDesc.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $tDesc.ScrollBars = "Vertical"; $tDesc.BorderStyle = "FixedSingle"
    $pnlGeral.Controls.Add($tDesc); $yy += 60

    # Status
    $lStatus = New-Object System.Windows.Forms.Label; $lStatus.Text = "Status:"; $lStatus.AutoSize = $true
    $lStatus.Location = New-Object System.Drawing.Point(15, $yy); $pnlGeral.Controls.Add($lStatus); $yy += 22
    $cmbStatus = New-Object System.Windows.Forms.ComboBox
    $cmbStatus.Items.AddRange(@("Habilitada (tudo ativo)", "Config de Usuario Desabilitada", "Config de Computador Desabilitada", "Toda Desabilitada"))
    $cmbStatus.DropDownStyle = "DropDownList"
    $cmbStatus.Location = New-Object System.Drawing.Point(15, $yy); $cmbStatus.Size = New-Object System.Drawing.Size(400, 30)
    $cmbStatus.BackColor = $BgField; $cmbStatus.ForeColor = $FgText; $cmbStatus.Font = New-Object System.Drawing.Font("Segoe UI", 11)
    $curFlags = [int]$gpoEntry.flags.Value
    if ($curFlags -ge 0 -and $curFlags -le 3) { $cmbStatus.SelectedIndex = $curFlags } else { $cmbStatus.SelectedIndex = 0 }
    $pnlGeral.Controls.Add($cmbStatus); $yy += 45

    # Separador
    $sep1 = New-Object System.Windows.Forms.Label; $sep1.AutoSize = $false; $sep1.Height = 2
    $sep1.Width = 900; $sep1.BackColor = $BgField; $sep1.Location = New-Object System.Drawing.Point(15, $yy)
    $pnlGeral.Controls.Add($sep1); $yy += 15

    # SECAO: Detalhes Tecnicose
    $sh2 = New-Object System.Windows.Forms.Label; $sh2.Text = "DETALHES TECNICOS"; $sh2.AutoSize = $true
    $sh2.ForeColor = $Accent; $sh2.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $sh2.Location = New-Object System.Drawing.Point(15, $yy); $pnlGeral.Controls.Add($sh2); $yy += 28

    # Grid de detalhes
    $dgvInfo = New-Object System.Windows.Forms.DataGridView
    $dgvInfo.Location = New-Object System.Drawing.Point(15, $yy)
    $dgvInfo.Size = New-Object System.Drawing.Size(900, 250)
    $dgvInfo.BackgroundColor = $BgDark; $dgvInfo.ForeColor = $FgText; $dgvInfo.GridColor = $BgField
    $dgvInfo.BorderStyle = "None"; $dgvInfo.ColumnHeadersVisible = $false
    $dgvInfo.RowHeadersVisible = $false; $dgvInfo.AllowUserToAddRows = $false
    $dgvInfo.ReadOnly = $true; $dgvInfo.SelectionMode = "FullRowSelect"
    $dgvInfo.DefaultCellStyle.BackColor = $BgDark; $dgvInfo.DefaultCellStyle.ForeColor = $FgText
    $dgvInfo.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(60,60,90)
    $dgvInfo.AutoSizeColumnsMode = "Fill"; $dgvInfo.EnableHeadersVisualStyles = $false
    $dgvInfo.Columns.Add("Prop", "Propriedade") | Out-Null
    $dgvInfo.Columns.Add("Val", "Valor") | Out-Null
    $dgvInfo.Columns["Prop"].FillWeight = 25; $dgvInfo.Columns["Val"].FillWeight = 75
    $dgvInfo.Columns["Prop"].DefaultCellStyle.ForeColor = $Accent
    $dgvInfo.Columns["Prop"].DefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    # Coletar TODAS as propriedades
    $gpoVersion = if ($gpoEntry.versionNumber) { [string]$gpoEntry.versionNumber } else { "0" }
    $gpoUserVer = [math]::Floor([int]$gpoVersion / 65536)
    $gpoCompVer = [int]$gpoVersion % 65536
    $gpoWhenCreated = [string]$gpoEntry.whenCreated
    $gpoWhenChanged = [string]$gpoEntry.whenChanged
    $gpoDN = [string]$gpoEntry.distinguishedName
    $gpoFileSys = [string]$gpoEntry.gPCFileSysPath
    $gpoFuncVer = if ($gpoEntry.gPCFunctionalityVersion) { [string]$gpoEntry.gPCFunctionalityVersion } else { "" }
    $gpoUserExt = if ($gpoEntry.gPCUserExtensionNames) { [string]$gpoEntry.gPCUserExtensionNames } else { "(nenhuma)" }
    $gpoMachExt = if ($gpoEntry.gPCMachineExtensionNames) { [string]$gpoEntry.gPCMachineExtensionNames } else { "(nenhuma)" }

    $flagsDesc = switch ($curFlags) { 0 {"Habilitada"} 1 {"User Desab."} 2 {"Computer Desab."} 3 {"Toda Desab."} default {"$curFlags"} }

    $infoRows = @(
        @("GUID", $GpoId),
        @("Distinguished Name", $gpoDN),
        @("Caminho SYSVOL", $gpoFileSys),
        @("Flags (Status)", "$curFlags - $flagsDesc"),
        @("Versao Total", $gpoVersion),
        @("Versao Usuario", "$gpoUserVer"),
        @("Versao Computador", "$gpoCompVer"),
        @("Functionality Version", $gpoFuncVer),
        @("Data de Criacao", $gpoWhenCreated),
        @("Data de Modificacao", $gpoWhenChanged),
        @("Extensoes de Usuario (CSEs)", $gpoUserExt),
        @("Extensoes de Computador (CSEs)", $gpoMachExt),
        @("Dominio", $script:Domain),
        @("Dominio DN", $domDN)
    )
    foreach ($row in $infoRows) { $dgvInfo.Rows.Add($row[0], $row[1]) | Out-Null }

    $pnlGeral.Controls.Add($dgvInfo)

    $tabGeral.Controls.Add($pnlGeral)
    $tabs.TabPages.Add($tabGeral)

    # ════════════════════════════════════════
    #  TAB 2: APLICAR A (ARVORE DO AD COM CHECKBOXES)
    # ════════════════════════════════════════
    $tabTarget = New-Object System.Windows.Forms.TabPage
    $tabTarget.Text = "  Aplicar a (AD)  "
    $tabTarget.BackColor = $BgDark

    $pnlTarget = New-Object System.Windows.Forms.Panel
    $pnlTarget.Dock = "Fill"

    # Header
    $pnlTargetHdr = New-Object System.Windows.Forms.Panel
    $pnlTargetHdr.Dock = "Top"; $pnlTargetHdr.Height = 60; $pnlTargetHdr.BackColor = $BgPanel

    $lblTargetTitle = New-Object System.Windows.Forms.Label
    $lblTargetTitle.Text = "ESTRUTURA DO AD - Marque as OUs onde esta GPO sera aplicada"
    $lblTargetTitle.AutoSize = $true; $lblTargetTitle.ForeColor = $Yellow
    $lblTargetTitle.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $lblTargetTitle.Location = New-Object System.Drawing.Point(15, 6)
    $pnlTargetHdr.Controls.Add($lblTargetTitle)

    $lblTargetSub = New-Object System.Windows.Forms.Label
    $lblTargetSub.Text = "Caixas MARCADAS = GPO vinculada.  Marque/desmarque para vincular/desvincular.  Expanda as pastas para ver usuarios e computadores."
    $lblTargetSub.AutoSize = $true; $lblTargetSub.ForeColor = [System.Drawing.Color]::FromArgb(108,112,134)
    $lblTargetSub.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblTargetSub.Location = New-Object System.Drawing.Point(15, 32)
    $pnlTargetHdr.Controls.Add($lblTargetSub)
    $pnlTarget.Controls.Add($pnlTargetHdr)

    # Painel de resumo dos vinculos ativos
    $pnlLinkedSummary = New-Object System.Windows.Forms.Panel
    $pnlLinkedSummary.Dock = "Top"; $pnlLinkedSummary.Height = 50; $pnlLinkedSummary.BackColor = [System.Drawing.Color]::FromArgb(25, 25, 40)
    $pnlLinkedSummary.Padding = New-Object System.Windows.Forms.Padding(15, 4, 15, 4)

    $lblLinkedIcon = New-Object System.Windows.Forms.Label
    $lblLinkedIcon.Text = "VINCULADA EM:"; $lblLinkedIcon.AutoSize = $true
    $lblLinkedIcon.ForeColor = $Green; $lblLinkedIcon.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblLinkedIcon.Location = New-Object System.Drawing.Point(15, 5)
    $pnlLinkedSummary.Controls.Add($lblLinkedIcon)

    $lblLinkedList = New-Object System.Windows.Forms.Label
    $lblLinkedList.Text = "(carregando...)"; $lblLinkedList.AutoSize = $false
    $lblLinkedList.Size = New-Object System.Drawing.Size(1200, 36)
    $lblLinkedList.ForeColor = $Green; $lblLinkedList.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblLinkedList.Location = New-Object System.Drawing.Point(130, 5)
    $pnlLinkedSummary.Controls.Add($lblLinkedList)

    $pnlTarget.Controls.Add($pnlLinkedSummary)

    function Update-LinkedSummary {
        if ($script:linkedOUs.Count -eq 0) {
            $lblLinkedList.Text = "(nenhum vinculo ativo)"
            $lblLinkedList.ForeColor = $Overlay
            $lblLinkedIcon.ForeColor = $Overlay
        } else {
            $names = @()
            foreach ($dn in $script:linkedOUs) {
                $parts = $dn -split ","
                $ouName = ($parts[0] -split "=")[1]
                $names += $ouName
            }
            $lblLinkedList.Text = $names -join "   |   "
            $lblLinkedList.ForeColor = $Green
            $lblLinkedIcon.ForeColor = $Green
        }
    }

    # Toolbar: filtro + legenda + botao aplicar
    $pnlADTool = New-Object System.Windows.Forms.Panel
    $pnlADTool.Dock = "Top"; $pnlADTool.Height = 42; $pnlADTool.BackColor = [System.Drawing.Color]::FromArgb(40,40,58)

    $lblFilterAD = New-Object System.Windows.Forms.Label; $lblFilterAD.Text = "Buscar:"; $lblFilterAD.AutoSize = $true
    $lblFilterAD.Location = New-Object System.Drawing.Point(10, 11); $lblFilterAD.ForeColor = $FgText
    $pnlADTool.Controls.Add($lblFilterAD)
    $txtFilterAD = New-Object System.Windows.Forms.TextBox
    $txtFilterAD.Location = New-Object System.Drawing.Point(65, 9); $txtFilterAD.Size = New-Object System.Drawing.Size(250, 24)
    $txtFilterAD.BackColor = $BgField; $txtFilterAD.ForeColor = $FgText
    $pnlADTool.Controls.Add($txtFilterAD)

    $btnExpandAll = New-Object System.Windows.Forms.Button
    $btnExpandAll.Text = "Expandir Tudo"; $btnExpandAll.Location = New-Object System.Drawing.Point(330, 7)
    $btnExpandAll.Size = New-Object System.Drawing.Size(110, 28); $btnExpandAll.BackColor = $BgField; $btnExpandAll.ForeColor = $FgText
    $btnExpandAll.FlatStyle = "Flat"; $btnExpandAll.Cursor = [System.Windows.Forms.Cursors]::Hand
    $pnlADTool.Controls.Add($btnExpandAll)

    $btnCollapseAll = New-Object System.Windows.Forms.Button
    $btnCollapseAll.Text = "Recolher Tudo"; $btnCollapseAll.Location = New-Object System.Drawing.Point(450, 7)
    $btnCollapseAll.Size = New-Object System.Drawing.Size(110, 28); $btnCollapseAll.BackColor = $BgField; $btnCollapseAll.ForeColor = $FgText
    $btnCollapseAll.FlatStyle = "Flat"; $btnCollapseAll.Cursor = [System.Windows.Forms.Cursors]::Hand
    $pnlADTool.Controls.Add($btnCollapseAll)

    $btnApplyLinks = New-Object System.Windows.Forms.Button
    $btnApplyLinks.Text = "APLICAR VINCULOS"; $btnApplyLinks.Location = New-Object System.Drawing.Point(580, 5)
    $btnApplyLinks.Size = New-Object System.Drawing.Size(170, 32); $btnApplyLinks.BackColor = $Green; $btnApplyLinks.ForeColor = $BgDark
    $btnApplyLinks.FlatStyle = "Flat"; $btnApplyLinks.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnApplyLinks.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $pnlADTool.Controls.Add($btnApplyLinks)

    # Legenda
    $lblLeg = New-Object System.Windows.Forms.Label
    $lblLeg.Text = "Pasta = OU  |  Pessoa = Usuario  |  PC = Computador  |  Grupo = Grupo"
    $lblLeg.AutoSize = $true; $lblLeg.ForeColor = [System.Drawing.Color]::FromArgb(108,112,134)
    $lblLeg.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $lblLeg.Location = New-Object System.Drawing.Point(770, 14)
    $pnlADTool.Controls.Add($lblLeg)

    $pnlTarget.Controls.Add($pnlADTool)

    # SplitContainer: Arvore esquerda + Detalhes direita
    $splitAD = New-Object System.Windows.Forms.SplitContainer
    $splitAD.Dock = "Fill"; $splitAD.Orientation = "Vertical"
    $splitAD.SplitterDistance = 550; $splitAD.BackColor = $BgField; $splitAD.SplitterWidth = 4

    # ── TreeView com checkboxes ──
    $treeAD = New-Object System.Windows.Forms.TreeView
    $treeAD.Dock = "Fill"
    $treeAD.CheckBoxes = $true
    $treeAD.BackColor = $BgDark; $treeAD.ForeColor = $FgText
    $treeAD.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $treeAD.LineColor = $BgField
    $treeAD.BorderStyle = "None"
    $treeAD.ShowPlusMinus = $true; $treeAD.ShowLines = $true; $treeAD.ShowRootLines = $true
    $treeAD.FullRowSelect = $true
    $treeAD.ItemHeight = 26

    # Buscar links atuais da GPO - ler gPLink de cada OU/container/dominio
    $script:linkedOUs = [System.Collections.ArrayList]@()
    $script:gpoIdUpper = $GpoId.ToUpper()

    # Verificar se o gPLink de um DN contem nosso GUID
    function Test-GPOLinked {
        param([string]$gpLinkValue)
        if ([string]::IsNullOrEmpty($gpLinkValue)) { return $false }
        return $gpLinkValue.ToUpper().Contains($script:gpoIdUpper)
    }

    # Checar dominio raiz
    try {
        $rootEntry = [ADSI]"LDAP://$domDN"
        $rootGPLink = [string]$rootEntry.Properties["gPLink"].Value
        if (Test-GPOLinked $rootGPLink) { $script:linkedOUs.Add($domDN) | Out-Null }
    } catch {}

    # Construir arvore do AD
    function Build-ADTree {
        $dlg.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        [System.Windows.Forms.Application]::DoEvents()
        $treeAD.BeginUpdate()
        $treeAD.Nodes.Clear()
        $script:linkedOUs.Clear()

        # Checar dominio raiz
        try {
            $rootEntry = [ADSI]"LDAP://$domDN"
            $rootGPLink = [string]$rootEntry.Properties["gPLink"].Value
            if (Test-GPOLinked $rootGPLink) { $script:linkedOUs.Add($domDN) | Out-Null }
        } catch {}

        # Raiz do dominio
        $rootNode = New-Object System.Windows.Forms.TreeNode
        $rootNode.Text = "$($script:Domain)  [$domDN]"
        $rootNode.Tag = @{ Type="OU"; DN=$domDN }
        $rootNode.NodeFont = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        if ($script:linkedOUs -contains $domDN) { $rootNode.Checked = $true; $rootNode.ForeColor = $Green }
        $treeAD.Nodes.Add($rootNode) | Out-Null

        # Buscar TUDO: OUs, Users, Computers, Groups
        $allObjects = [System.Collections.ArrayList]@()

        # OUs - incluir gPLink para detectar vinculos
        try {
            $s = New-Object System.DirectoryServices.DirectorySearcher
            $s.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$domDN")
            $s.Filter = "(objectClass=organizationalUnit)"; $s.PageSize = 1000; $s.SearchScope = "Subtree"
            $s.PropertiesToLoad.AddRange(@("distinguishedName","name","description","gPLink")) | Out-Null
            $res = $s.FindAll()
            foreach ($r in $res) {
                $dn = [string]$r.Properties["distinguishedname"][0]
                $gpLink = if ($r.Properties["gplink"].Count -gt 0) { [string]$r.Properties["gplink"][0] } else { "" }
                $isLinked = Test-GPOLinked $gpLink
                if ($isLinked -and ($script:linkedOUs -notcontains $dn)) { $script:linkedOUs.Add($dn) | Out-Null }
                $allObjects.Add(@{
                    Type = "OU"
                    DN   = $dn
                    Name = [string]$r.Properties["name"][0]
                    Desc = if ($r.Properties["description"].Count -gt 0) { [string]$r.Properties["description"][0] } else { "" }
                    Extra = ""
                }) | Out-Null
            }
            $res.Dispose(); $s.Dispose()
        } catch {}

        # Containers (CN=Users, CN=Computers, CN=Builtin) - incluir gPLink
        try {
            $s = New-Object System.DirectoryServices.DirectorySearcher
            $s.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$domDN")
            $s.Filter = "(&(objectClass=container)(|(cn=Users)(cn=Computers)(cn=Builtin)))"; $s.PageSize = 100
            $s.PropertiesToLoad.AddRange(@("distinguishedName","cn","gPLink")) | Out-Null
            $res = $s.FindAll()
            foreach ($r in $res) {
                $dn = [string]$r.Properties["distinguishedname"][0]
                $gpLink = if ($r.Properties["gplink"].Count -gt 0) { [string]$r.Properties["gplink"][0] } else { "" }
                $isLinked = Test-GPOLinked $gpLink
                if ($isLinked -and ($script:linkedOUs -notcontains $dn)) { $script:linkedOUs.Add($dn) | Out-Null }
                $allObjects.Add(@{
                    Type = "OU"
                    DN   = $dn
                    Name = [string]$r.Properties["cn"][0]
                    Desc = "(Container padrao)"
                    Extra = ""
                }) | Out-Null
            }
            $res.Dispose(); $s.Dispose()
        } catch {}

        # Usuarios
        try {
            $s = New-Object System.DirectoryServices.DirectorySearcher
            $s.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$domDN")
            $s.Filter = "(&(objectClass=user)(objectCategory=person))"; $s.PageSize = 1000
            $s.PropertiesToLoad.AddRange(@("distinguishedName","cn","sAMAccountName","department","title","userAccountControl","mail")) | Out-Null
            $res = $s.FindAll()
            foreach ($r in $res) {
                $uac = if ($r.Properties["useraccountcontrol"].Count -gt 0) { [int]$r.Properties["useraccountcontrol"][0] } else { 0 }
                $enabled = -not ($uac -band 2)
                $login = if ($r.Properties["samaccountname"].Count -gt 0) { [string]$r.Properties["samaccountname"][0] } else { "" }
                $dept = if ($r.Properties["department"].Count -gt 0) { [string]$r.Properties["department"][0] } else { "" }
                $title = if ($r.Properties["title"].Count -gt 0) { [string]$r.Properties["title"][0] } else { "" }
                $mail = if ($r.Properties["mail"].Count -gt 0) { [string]$r.Properties["mail"][0] } else { "" }
                $statusTxt = if ($enabled) { "Ativo" } else { "Desativado" }
                $allObjects.Add(@{
                    Type = "User"
                    DN   = [string]$r.Properties["distinguishedname"][0]
                    Name = [string]$r.Properties["cn"][0]
                    Desc = "$login | $dept | $title | $mail | $statusTxt"
                    Extra = $statusTxt
                    Enabled = $enabled
                }) | Out-Null
            }
            $res.Dispose(); $s.Dispose()
        } catch {}

        # Computadores
        try {
            $s = New-Object System.DirectoryServices.DirectorySearcher
            $s.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$domDN")
            $s.Filter = "(objectClass=computer)"; $s.PageSize = 1000
            $s.PropertiesToLoad.AddRange(@("distinguishedName","cn","operatingSystem","description")) | Out-Null
            $res = $s.FindAll()
            foreach ($r in $res) {
                $os = if ($r.Properties["operatingsystem"].Count -gt 0) { [string]$r.Properties["operatingsystem"][0] } else { "" }
                $desc = if ($r.Properties["description"].Count -gt 0) { [string]$r.Properties["description"][0] } else { "" }
                $allObjects.Add(@{
                    Type = "Computer"
                    DN   = [string]$r.Properties["distinguishedname"][0]
                    Name = [string]$r.Properties["cn"][0]
                    Desc = "$os | $desc"
                    Extra = $os
                }) | Out-Null
            }
            $res.Dispose(); $s.Dispose()
        } catch {}

        # Grupos
        try {
            $s = New-Object System.DirectoryServices.DirectorySearcher
            $s.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$domDN")
            $s.Filter = "(objectClass=group)"; $s.PageSize = 1000
            $s.PropertiesToLoad.AddRange(@("distinguishedName","cn","description","groupType","member")) | Out-Null
            $res = $s.FindAll()
            foreach ($r in $res) {
                $gt = if ($r.Properties["grouptype"].Count -gt 0) { [int]$r.Properties["grouptype"][0] } else { 0 }
                $tipo = if ($gt -band 0x80000000) { "Seguranca" } else { "Distribuicao" }
                $mc = $r.Properties["member"].Count
                $desc = if ($r.Properties["description"].Count -gt 0) { [string]$r.Properties["description"][0] } else { "" }
                $allObjects.Add(@{
                    Type = "Group"
                    DN   = [string]$r.Properties["distinguishedname"][0]
                    Name = [string]$r.Properties["cn"][0]
                    Desc = "$tipo | $mc membros | $desc"
                    Extra = $tipo
                    Members = $mc
                }) | Out-Null
            }
            $res.Dispose(); $s.Dispose()
        } catch {}

        # Montar hash de nodes por DN para lookup rapido
        $nodeMap = @{}
        $nodeMap[$domDN] = $rootNode

        # Primeiro: criar nodes de OUs/Containers
        $ous = $allObjects | Where-Object { $_.Type -eq "OU" } | Sort-Object { $_.DN.Split(',').Count }
        foreach ($ou in $ous) {
            $parentDN = ($ou.DN -split ',', 2)[1]
            $parentNode = if ($nodeMap.ContainsKey($parentDN)) { $nodeMap[$parentDN] } else { $rootNode }

            $ouNode = New-Object System.Windows.Forms.TreeNode
            $ouNode.Text = "[OU] $($ou.Name)"
            $ouNode.Tag = @{ Type="OU"; DN=$ou.DN; Desc=$ou.Desc }
            $ouNode.NodeFont = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
            if ($script:linkedOUs -contains $ou.DN) {
                $ouNode.Checked = $true
                $ouNode.ForeColor = $Green
            }
            $parentNode.Nodes.Add($ouNode) | Out-Null
            $nodeMap[$ou.DN] = $ouNode
        }

        # Depois: adicionar users, computers, groups dentro das OUs
        $nonOUs = $allObjects | Where-Object { $_.Type -ne "OU" }
        foreach ($obj in $nonOUs) {
            $parentDN = ($obj.DN -split ',', 2)[1]
            $parentNode = if ($nodeMap.ContainsKey($parentDN)) { $nodeMap[$parentDN] } else { $rootNode }

            $childNode = New-Object System.Windows.Forms.TreeNode
            $prefix = switch ($obj.Type) { "User" { "👤" } "Computer" { "💻" } "Group" { "👥" } default { "•" } }
            $childNode.Text = "$prefix $($obj.Name)"
            $childNode.Tag = @{ Type=$obj.Type; DN=$obj.DN; Desc=$obj.Desc }

            if ($obj.Type -eq "User") {
                if ($obj.Enabled -eq $false) {
                    $childNode.ForeColor = $Red
                } else {
                    $childNode.ForeColor = $FgText
                }
            } elseif ($obj.Type -eq "Computer") {
                $childNode.ForeColor = $Accent
            } elseif ($obj.Type -eq "Group") {
                $childNode.ForeColor = $Yellow
            }

            $parentNode.Nodes.Add($childNode) | Out-Null
        }

        $rootNode.Expand()
        # Expandir OUs que estao linkadas
        foreach ($key in $nodeMap.Keys) {
            if ($script:linkedOUs -contains $key) { $nodeMap[$key].Expand() }
        }

        $treeAD.EndUpdate()
        if ($treeAD.Nodes.Count -gt 0) { $treeAD.Nodes[0].EnsureVisible() }
        $dlg.Cursor = [System.Windows.Forms.Cursors]::Default
    }

    Build-ADTree
    Update-LinkedSummary

    # Botoes
    $btnExpandAll.Add_Click({ $treeAD.ExpandAll() })
    $btnCollapseAll.Add_Click({ $treeAD.CollapseAll() })

    # Filtro de busca na arvore
    $txtFilterAD.Add_TextChanged({
        $filter = $txtFilterAD.Text.Trim().ToLower()
        if ([string]::IsNullOrEmpty($filter)) {
            # Sem filtro: reconstruir tudo
            Build-ADTree
            return
        }
        # Com filtro: destacar nodes que contem o texto
        $treeAD.BeginUpdate()
        foreach ($rootN in $treeAD.Nodes) {
            Filter-TreeNode $rootN $filter
        }
        $treeAD.EndUpdate()
    })

    function Filter-TreeNode {
        param($node, $filter)
        $match = $node.Text.ToLower().Contains($filter)
        $childMatch = $false
        foreach ($child in $node.Nodes) {
            if (Filter-TreeNode $child $filter) { $childMatch = $true }
        }
        if ($match -or $childMatch) {
            if ($match) { $node.BackColor = [System.Drawing.Color]::FromArgb(60,60,90) }
            else { $node.BackColor = $BgDark }
            $node.Expand()
            return $true
        } else {
            $node.BackColor = $BgDark
            return $false
        }
    }

    # Painel de detalhes (direita)
    $pnlDetail = New-Object System.Windows.Forms.Panel
    $pnlDetail.Dock = "Fill"; $pnlDetail.BackColor = $BgPanel; $pnlDetail.AutoScroll = $true
    $pnlDetail.Padding = New-Object System.Windows.Forms.Padding(15)

    $lblDetailTitle = New-Object System.Windows.Forms.Label
    $lblDetailTitle.Text = "DETALHES"; $lblDetailTitle.AutoSize = $true
    $lblDetailTitle.ForeColor = $Accent; $lblDetailTitle.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $lblDetailTitle.Location = New-Object System.Drawing.Point(15, 10)
    $pnlDetail.Controls.Add($lblDetailTitle)

    $lblDetailInfo = New-Object System.Windows.Forms.Label
    $lblDetailInfo.Text = "Selecione um item na arvore para ver detalhes."
    $lblDetailInfo.AutoSize = $false; $lblDetailInfo.Size = New-Object System.Drawing.Size(380, 400)
    $lblDetailInfo.ForeColor = $FgText; $lblDetailInfo.Font = New-Object System.Drawing.Font("Consolas", 9)
    $lblDetailInfo.Location = New-Object System.Drawing.Point(15, 40)
    $pnlDetail.Controls.Add($lblDetailInfo)

    # Ao selecionar um node, mostrar detalhes
    $treeAD.Add_AfterSelect({
        param($s4, $e4)
        $node = $e4.Node
        if (-not $node.Tag) { return }
        $info = $node.Tag
        $text = "Tipo: $($info.Type)`n"
        $text += "DN: $($info.DN)`n`n"
        if ($info.Desc) { $text += "Detalhes:`n$($info.Desc)`n`n" }

        if ($info.Type -eq "OU") {
            $isLinked = if ($script:linkedOUs -contains $info.DN) { "SIM - GPO vinculada aqui" } else { "NAO" }
            $text += "GPO Vinculada: $isLinked`n"

            # Contar filhos
            $totalChildren = $node.Nodes.Count
            $userCount = 0; $compCount = 0; $grpCount = 0; $ouCount = 0
            foreach ($child in $node.Nodes) {
                if ($child.Tag) {
                    switch ($child.Tag.Type) {
                        "User"     { $userCount++ }
                        "Computer" { $compCount++ }
                        "Group"    { $grpCount++ }
                        "OU"       { $ouCount++ }
                    }
                }
            }
            $text += "`nConteudo desta OU:`n"
            $text += "  Sub-OUs:       $ouCount`n"
            $text += "  Usuarios:      $userCount`n"
            $text += "  Computadores:  $compCount`n"
            $text += "  Grupos:        $grpCount`n"
            $text += "  Total:         $totalChildren`n"
        }

        if ($info.Type -eq "User") {
            $parts = $info.Desc -split '\|'
            if ($parts.Count -ge 5) {
                $text += "Login:        $($parts[0].Trim())`n"
                $text += "Departamento: $($parts[1].Trim())`n"
                $text += "Cargo:        $($parts[2].Trim())`n"
                $text += "Email:        $($parts[3].Trim())`n"
                $text += "Status:       $($parts[4].Trim())`n"
            }
        }

        if ($info.Type -eq "Computer") {
            $parts = $info.Desc -split '\|'
            if ($parts.Count -ge 1) {
                $text += "Sistema:  $($parts[0].Trim())`n"
                if ($parts.Count -ge 2) { $text += "Descricao: $($parts[1].Trim())`n" }
            }
        }

        if ($info.Type -eq "Group") {
            $parts = $info.Desc -split '\|'
            if ($parts.Count -ge 2) {
                $text += "Tipo:     $($parts[0].Trim())`n"
                $text += "Membros:  $($parts[1].Trim())`n"
                if ($parts.Count -ge 3) { $text += "Descricao: $($parts[2].Trim())`n" }
            }
        }

        $lblDetailInfo.Text = $text
    })

    # APLICAR VINCULOS - quando clicar o botao verde
    $btnApplyLinks.Add_Click({
        # Coletar quais OUs estao marcadas
        $checkedOUs = [System.Collections.ArrayList]@()
        function Collect-CheckedOUs {
            param($nodes)
            foreach ($n in $nodes) {
                if ($n.Tag -and $n.Tag.Type -eq "OU" -and $n.Checked) {
                    $checkedOUs.Add($n.Tag.DN) | Out-Null
                }
                if ($n.Nodes.Count -gt 0) { Collect-CheckedOUs $n.Nodes }
            }
        }
        Collect-CheckedOUs $treeAD.Nodes

        # Comparar com estado atual
        $toLink   = $checkedOUs | Where-Object { $_ -notin $script:linkedOUs }
        $toUnlink = $script:linkedOUs | Where-Object { $_ -notin $checkedOUs }

        if ($toLink.Count -eq 0 -and $toUnlink.Count -eq 0) {
            Show-DarkMsg "Nenhuma alteracao de vinculo." "Info" "OK" "Information"
            return
        }

        $msg = ""
        if ($toLink.Count -gt 0) { $msg += "VINCULAR a:`n" + ($toLink | ForEach-Object { "  + $_" }) -join "`n"; $msg += "`n`n" }
        if ($toUnlink.Count -gt 0) { $msg += "DESVINCULAR de:`n" + ($toUnlink | ForEach-Object { "  - $_" }) -join "`n" }

        $conf = Show-DarkMsg "Confirmar alteracoes de vinculo?`n`n$msg" "Aplicar Vinculos" "YesNo" "Question"
        if ($conf -ne "Yes") { return }

        $dlg.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $errors = @()

        # Vincular novas
        foreach ($ouDN in $toLink) {
            try {
                $ouObj = [ADSI]"LDAP://$ouDN"
                $curLink = [string]$ouObj.gPLink
                $newLinkEntry = "[LDAP://CN=$GpoId,CN=Policies,CN=System,$domDN;0]"
                if ($curLink -and $curLink -like "*$GpoId*") { continue }
                if ($curLink) { $ouObj.Put("gPLink", "$curLink$newLinkEntry") }
                else { $ouObj.Put("gPLink", $newLinkEntry) }
                $ouObj.SetInfo()
                if ($ouDN -notin $script:linkedOUs) { $script:linkedOUs.Add($ouDN) | Out-Null }
            } catch {
                $errors += "Vincular $ouDN : $($_.Exception.Message)"
            }
        }

        # Desvincular
        foreach ($ouDN in $toUnlink) {
            try {
                $ouObj = [ADSI]"LDAP://$ouDN"
                $curLink = [string]$ouObj.gPLink
                $escapedDomDN = [Regex]::Escape($domDN)
                $pattern = "\[LDAP://[Cc][Nn]=$([Regex]::Escape($GpoId)),CN=Policies,CN=System,$escapedDomDN;\d\]"
                $newLink = [regex]::Replace($curLink, $pattern, "", "IgnoreCase")
                if ($newLink.Trim() -eq "") { $ouObj.PutEx(1, "gPLink", $null) }
                else { $ouObj.Put("gPLink", $newLink) }
                $ouObj.SetInfo()
                $script:linkedOUs.Remove($ouDN)
            } catch {
                $errors += "Desvincular $ouDN : $($_.Exception.Message)"
            }
        }

        $dlg.Cursor = [System.Windows.Forms.Cursors]::Default

        if ($errors.Count -gt 0) {
            Show-DarkMsg "Concluido com erros:`n`n$($errors -join "`n")" "Aviso" "OK" "Warning"
        } else {
            Show-DarkMsg "Vinculos aplicados com sucesso!`n`nVinculados: $($toLink.Count)`nDesvinculados: $($toUnlink.Count)" "Sucesso" "OK" "Information"
        }

        # Rebuild tree
        Build-ADTree
        Update-LinkedSummary
    })

    $splitAD.Panel1.Controls.Add($treeAD)
    $splitAD.Panel2.Controls.Add($pnlDetail)
    $pnlTarget.Controls.Add($splitAD)

    # Corrigir ordem visual: WinForms Dock=Top empilha em ordem inversa
    # Ordem desejada (de cima pra baixo): Header -> LinkedSummary -> Toolbar -> SplitAD(Fill)
    $pnlTargetHdr.BringToFront()
    $pnlLinkedSummary.BringToFront()
    $pnlADTool.BringToFront()
    $splitAD.BringToFront()

    $tabTarget.Controls.Add($pnlTarget)
    $tabs.TabPages.Add($tabTarget)

    # ════════════════════════════════════════
    #  TAB 3: POLITICAS COMUNS (checkboxes)
    # ════════════════════════════════════════
    $tabPolicies = New-Object System.Windows.Forms.TabPage
    $tabPolicies.Text = "  Politicas Comuns  "
    $tabPolicies.BackColor = $BgDark

    # Resumo de politicas ativas no topo
    $pnlPolSummary = New-Object System.Windows.Forms.Panel
    $pnlPolSummary.Dock = "Top"; $pnlPolSummary.Height = 36; $pnlPolSummary.BackColor = [System.Drawing.Color]::FromArgb(25, 25, 40)
    $activeCount = $script:savedPolicies.Count
    $lblPolSummary = New-Object System.Windows.Forms.Label
    $lblPolSummary.AutoSize = $true; $lblPolSummary.Location = New-Object System.Drawing.Point(15, 9)
    $lblPolSummary.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    if ($activeCount -gt 0) {
        $activeNames = @()
        foreach ($pol in $commonPolicies) {
            if ($script:savedPolicies.ContainsKey($pol.Id)) { $activeNames += $pol.Name }
        }
        $lblPolSummary.Text = "ATIVAS ($activeCount):  $($activeNames -join '  |  ')"
        $lblPolSummary.ForeColor = $Green
    } else {
        $lblPolSummary.Text = "NENHUMA POLITICA ATIVA"
        $lblPolSummary.ForeColor = $Overlay
    }
    $pnlPolSummary.Controls.Add($lblPolSummary)
    $tabPolicies.Controls.Add($pnlPolSummary)

    $pnlPol = New-Object System.Windows.Forms.Panel
    $pnlPol.Dock = "Fill"; $pnlPol.AutoScroll = $true; $pnlPol.Padding = New-Object System.Windows.Forms.Padding(20)

    # Definir politicas comuns como array de objetos
    $commonPolicies = @(
        @{ Id="DisableControlPanel";  Cat="Restricoes do Sistema"; Name="Desabilitar Painel de Controle e Configuracoes"; Desc="Impede acesso ao Painel de Controle"; Hive="HKCU"; Key="SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer"; Val="NoControlPanel"; Data="1" }
        @{ Id="DisableCMD";           Cat="Restricoes do Sistema"; Name="Desabilitar Prompt de Comando (CMD)"; Desc="Bloqueia execucao do cmd.exe"; Hive="HKCU"; Key="SOFTWARE\Policies\Microsoft\Windows\System"; Val="DisableCMD"; Data="1" }
        @{ Id="DisableTaskMgr";       Cat="Restricoes do Sistema"; Name="Desabilitar Gerenciador de Tarefas"; Desc="Remove acesso ao Ctrl+Shift+Esc"; Hive="HKCU"; Key="SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"; Val="DisableTaskMgr"; Data="1" }
        @{ Id="DisableRegistry";      Cat="Restricoes do Sistema"; Name="Desabilitar Editor de Registro (regedit)"; Desc="Impede acesso ao regedit.exe"; Hive="HKCU"; Key="SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"; Val="DisableRegistryTools"; Data="1" }
        @{ Id="NoRun";                Cat="Restricoes do Sistema"; Name="Remover 'Executar' do Menu Iniciar"; Desc="Esconde a opcao Executar (Win+R continua)"; Hive="HKCU"; Key="SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer"; Val="NoRun"; Data="1" }
        @{ Id="DisableChangePwd";     Cat="Restricoes do Sistema"; Name="Desabilitar alteracao de senha pelo usuario"; Desc="Impede o usuario de trocar a propria senha"; Hive="HKCU"; Key="SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"; Val="DisableChangePassword"; Data="1" }
        @{ Id="DisableLockWS";        Cat="Restricoes do Sistema"; Name="Remover opcao 'Bloquear Computador'"; Desc="Remove a opcao de bloquear tela"; Hive="HKCU"; Key="SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"; Val="DisableLockWorkstation"; Data="1" }

        @{ Id="NoWallpaper";          Cat="Personalizacao"; Name="Impedir alteracao do papel de parede"; Desc="Usuario nao pode mudar o wallpaper"; Hive="HKCU"; Key="SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop"; Val="NoChangingWallPaper"; Data="1" }
        @{ Id="NoThemes";             Cat="Personalizacao"; Name="Impedir alteracao de temas e cores"; Desc="Bloqueia personalizacao visual"; Hive="HKCU"; Key="SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer"; Val="NoThemesTab"; Data="1" }
        @{ Id="NoDispCPL";            Cat="Personalizacao"; Name="Desabilitar Configuracoes de Video"; Desc="Remove acesso a configuracao de tela"; Hive="HKCU"; Key="SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"; Val="NoDispCPL"; Data="1" }

        @{ Id="DisableUSB";           Cat="Dispositivos"; Name="Desabilitar armazenamento USB"; Desc="Bloqueia pen drives e discos USB"; Hive="HKLM"; Key="SYSTEM\CurrentControlSet\Services\USBSTOR"; Val="Start"; Data="4" }
        @{ Id="DisableCDROM";         Cat="Dispositivos"; Name="Desabilitar unidade de CD/DVD"; Desc="Bloqueia acesso ao drive de CD-ROM"; Hive="HKLM"; Key="SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer"; Val="NoDriveTypeAutoRun"; Data="255" }

        @{ Id="DisableWindowsUpdate"; Cat="Windows Update"; Name="Desabilitar Windows Update automatico"; Desc="Impede atualizacoes automaticas"; Hive="HKLM"; Key="SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"; Val="NoAutoUpdate"; Data="1" }
        @{ Id="NoAutoRestart";        Cat="Windows Update"; Name="Nao reiniciar automaticamente (com usuario logado)"; Desc="Impede restart apos update"; Hive="HKLM"; Key="SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"; Val="NoAutoRebootWithLoggedOnUsers"; Data="1" }

        @{ Id="DisableRemoteDesktop"; Cat="Rede e Acesso"; Name="Desabilitar Acesso Remoto (RDP)"; Desc="Bloqueia conexoes RDP ao computador"; Hive="HKLM"; Key="SYSTEM\CurrentControlSet\Control\Terminal Server"; Val="fDenyTSConnections"; Data="1" }
        @{ Id="EnableFirewall";       Cat="Rede e Acesso"; Name="Forcar Firewall do Windows ativado"; Desc="Garante que o firewall fique ligado"; Hive="HKLM"; Key="SOFTWARE\Policies\Microsoft\WindowsFirewall\DomainProfile"; Val="EnableFirewall"; Data="1" }
        @{ Id="DisableNetworkDiscover";Cat="Rede e Acesso"; Name="Desabilitar descoberta de rede"; Desc="Impede que o PC seja visivel na rede"; Hive="HKLM"; Key="SOFTWARE\Policies\Microsoft\Windows\LLTD"; Val="EnableLLTDIO"; Data="0" }

        @{ Id="NoPwdReveal";          Cat="Seguranca"; Name="Ocultar botao de revelar senha"; Desc="Remove o olho nos campos de senha"; Hive="HKLM"; Key="SOFTWARE\Policies\Microsoft\Windows\CredUI"; Val="DisablePasswordReveal"; Data="1" }
        @{ Id="MinPwdAge";            Cat="Seguranca"; Name="Exigir senha complexa (min 8 chars)"; Desc="Forca senha com pelo menos 8 caracteres"; Hive="HKLM"; Key="SOFTWARE\Policies\Microsoft\Windows\System"; Val="MinimumPasswordLength"; Data="8" }
        @{ Id="RequireCtrlAltDel";    Cat="Seguranca"; Name="Exigir Ctrl+Alt+Del para login"; Desc="Adiciona camada de seguranca no logon"; Hive="HKLM"; Key="SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"; Val="DisableCAD"; Data="0" }
        @{ Id="LockScreenTimeout";    Cat="Seguranca"; Name="Bloquear tela apos 10 min inativo"; Desc="Trava o PC automaticamente"; Hive="HKCU"; Key="SOFTWARE\Policies\Microsoft\Windows\Control Panel\Desktop"; Val="ScreenSaverIsSecure"; Data="1" }

        @{ Id="HideLogonName";        Cat="Login / Logon"; Name="Nao exibir ultimo usuario logado"; Desc="Nao mostra o nome do ultimo login na tela"; Hive="HKLM"; Key="SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"; Val="DontDisplayLastUserName"; Data="1" }
        @{ Id="LegalNotice";          Cat="Login / Logon"; Name="Exibir aviso legal antes do login"; Desc="Mostra mensagem obrigatoria ao ligar PC"; Hive="HKLM"; Key="SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"; Val="LegalNoticeCaption"; Data="AVISO" }
        @{ Id="DisableShutdown";      Cat="Login / Logon"; Name="Remover botao Desligar na tela de login"; Desc="Impede desligamento sem logar"; Hive="HKLM"; Key="SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"; Val="ShutdownWithoutLogon"; Data="0" }
        @{ Id="AutoLogoff";           Cat="Login / Logon"; Name="Desconectar quando exceder horario"; Desc="Forca logoff fora do horario"; Hive="HKLM"; Key="SYSTEM\CurrentControlSet\Services\LanManServer\Parameters"; Val="EnableForcedLogOff"; Data="1" }

        @{ Id="NoAutoPlay";           Cat="Midia / AutoPlay"; Name="Desabilitar AutoPlay em todas as unidades"; Desc="Nao executa automaticamente ao inserir midia"; Hive="HKLM"; Key="SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer"; Val="NoDriveTypeAutoRun"; Data="255" }
        @{ Id="DisableStore";         Cat="Windows Store"; Name="Desabilitar Microsoft Store"; Desc="Bloqueia acesso a loja de apps"; Hive="HKLM"; Key="SOFTWARE\Policies\Microsoft\WindowsStore"; Val="RemoveWindowsStore"; Data="1" }
        @{ Id="DisableCortana";       Cat="Assistente"; Name="Desabilitar Cortana"; Desc="Desliga assistente da Microsoft"; Hive="HKLM"; Key="SOFTWARE\Policies\Microsoft\Windows\Windows Search"; Val="AllowCortana"; Data="0" }
        @{ Id="DisableOneDrive";      Cat="Cloud"; Name="Desabilitar OneDrive"; Desc="Impede uso do OneDrive"; Hive="HKLM"; Key="SOFTWARE\Policies\Microsoft\Windows\OneDrive"; Val="DisableFileSyncNGSC"; Data="1" }
        @{ Id="DisableTelemetry";     Cat="Privacidade"; Name="Desabilitar telemetria do Windows"; Desc="Para de enviar dados a Microsoft"; Hive="HKLM"; Key="SOFTWARE\Policies\Microsoft\Windows\DataCollection"; Val="AllowTelemetry"; Data="0" }
        @{ Id="DisableErrorReport";   Cat="Privacidade"; Name="Desabilitar Relatorio de Erros"; Desc="Nao envia relatorios de crash"; Hive="HKLM"; Key="SOFTWARE\Policies\Microsoft\Windows\Windows Error Reporting"; Val="Disabled"; Data="1" }

        @{ Id="MapNetworkDrive";      Cat="Mapeamento de Rede"; Name="Reconectar unidades mapeadas no logon"; Desc="Restaura drives de rede ao logar"; Hive="HKCU"; Key="SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer"; Val="NoPersistentConnections"; Data="0" }
        @{ Id="HideNetworkDrives";    Cat="Mapeamento de Rede"; Name="Ocultar drives especificos no Explorer"; Desc="Esconde unidades de disco"; Hive="HKCU"; Key="SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer"; Val="NoDrives"; Data="0" }

        @{ Id="DisablePowerSettings";  Cat="Energia"; Name="Impedir alteracao de plano de energia"; Desc="Usuario nao muda config de energia"; Hive="HKLM"; Key="SOFTWARE\Policies\Microsoft\Power\PowerSettings"; Val="ActivePowerScheme"; Data="1" }
        @{ Id="NoSleep";              Cat="Energia"; Name="Desabilitar modo suspensao"; Desc="PC nunca entra em sleep"; Hive="HKLM"; Key="SOFTWARE\Policies\Microsoft\Power\PowerSettings\abfc2519-3608-4c2a-94ea-171b0ed546ab"; Val="ACSettingIndex"; Data="0" }

        @{ Id="AuditLogon";           Cat="Auditoria"; Name="Auditar eventos de logon"; Desc="Registra tentativas de login no Event Log"; Hive="HKLM"; Key="SOFTWARE\Policies\Microsoft\Windows\System"; Val="AuditLogon"; Data="3" }
        @{ Id="AuditObjectAccess";    Cat="Auditoria"; Name="Auditar acesso a objetos"; Desc="Registra acesso a arquivos e pastas"; Hive="HKLM"; Key="SOFTWARE\Policies\Microsoft\Windows\System"; Val="AuditObjectAccess"; Data="3" }
    )

    # ── Mapear Registry.pol nativo para politicas conhecidas ──
    foreach ($regEntry in $script:nativeRegEntries) {
        foreach ($pol in $commonPolicies) {
            if ($regEntry.Key -like "*$($pol.Key)*" -and $regEntry.ValueName -eq $pol.Val) {
                if (-not $script:savedPolicies.ContainsKey($pol.Id)) {
                    $script:savedPolicies[$pol.Id] = $true
                }
                break
            }
        }
    }

    # Atualizar resumo apos mapeamento
    $activeCount = $script:savedPolicies.Count
    if ($activeCount -gt 0) {
        $activeNames = @()
        foreach ($pol in $commonPolicies) {
            if ($script:savedPolicies.ContainsKey($pol.Id)) { $activeNames += $pol.Name }
        }
        $lblPolSummary.Text = "ATIVAS ($activeCount):  $($activeNames -join '  |  ')"
        $lblPolSummary.ForeColor = $Green
    } else {
        $lblPolSummary.Text = "NENHUMA POLITICA ATIVA"
        $lblPolSummary.ForeColor = $Overlay
    }

    $py = 15
    $lastCat = ""
    $chkBoxes = @{}
    foreach ($pol in $commonPolicies) {
        # Cabecalho de categoria
        if ($pol.Cat -ne $lastCat) {
            if ($lastCat -ne "") { $py += 8 }
            $catLbl = New-Object System.Windows.Forms.Label
            $catLbl.Text = $pol.Cat.ToUpper()
            $catLbl.AutoSize = $true; $catLbl.ForeColor = $Accent
            $catLbl.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
            $catLbl.Location = New-Object System.Drawing.Point(15, $py)
            $pnlPol.Controls.Add($catLbl)
            $py += 26
            # Linha
            $line = New-Object System.Windows.Forms.Label; $line.AutoSize = $false
            $line.Size = New-Object System.Drawing.Size(770, 1); $line.BackColor = $BgField
            $line.Location = New-Object System.Drawing.Point(15, $py); $pnlPol.Controls.Add($line); $py += 6
            $lastCat = $pol.Cat
        }

        # Checkbox
        $chk = New-Object System.Windows.Forms.CheckBox
        $chk.Text = "  $($pol.Name)"
        $chk.AutoSize = $true
        $chk.ForeColor = $FgText
        $chk.Font = New-Object System.Drawing.Font("Segoe UI", 10)
        $chk.Location = New-Object System.Drawing.Point(20, $py)
        $chk.Tag = $pol.Id
        # Verificar se estava salvo
        if ($script:savedPolicies.ContainsKey($pol.Id)) {
            $chk.Checked = $true
            $chk.ForeColor = $Green
        }
        $pnlPol.Controls.Add($chk); $py += 6

        # Descricao
        $descLbl = New-Object System.Windows.Forms.Label
        $descLbl.Text = $pol.Desc
        $descLbl.AutoSize = $true
        $descLbl.ForeColor = [System.Drawing.Color]::FromArgb(108,112,134)
        $descLbl.Font = New-Object System.Drawing.Font("Segoe UI", 8)
        $descLbl.Location = New-Object System.Drawing.Point(42, ($py + 16))
        $pnlPol.Controls.Add($descLbl)
        $py += 38

        $chkBoxes[$pol.Id] = $chk
    }

    $tabPolicies.Controls.Add($pnlPol)
    $tabs.TabPages.Add($tabPolicies)

    # ════════════════════════════════════════
    #  TAB 3: BLOQUEIO DE APPS
    # ════════════════════════════════════════
    $tabApps = New-Object System.Windows.Forms.TabPage
    $tabApps.Text = "  Bloqueio de Apps  "
    $tabApps.BackColor = $BgDark

    $pnlAppsMain = New-Object System.Windows.Forms.Panel
    $pnlAppsMain.Dock = "Fill"

    # Header da aba
    $pnlAppsHeader = New-Object System.Windows.Forms.Panel
    $pnlAppsHeader.Dock = "Top"; $pnlAppsHeader.Height = 50; $pnlAppsHeader.BackColor = $BgPanel

    $lblAppsInfo = New-Object System.Windows.Forms.Label
    $lblAppsInfo.Text = "Marque os programas que devem ser BLOQUEADOS nesta GPO:"
    $lblAppsInfo.AutoSize = $true; $lblAppsInfo.ForeColor = $Yellow
    $lblAppsInfo.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $lblAppsInfo.Location = New-Object System.Drawing.Point(15, 14)
    $pnlAppsHeader.Controls.Add($lblAppsInfo)
    $pnlAppsMain.Controls.Add($pnlAppsHeader)

    # Painel scrollavel com checkboxes de apps
    $pnlAppsScroll = New-Object System.Windows.Forms.Panel
    $pnlAppsScroll.Dock = "Fill"; $pnlAppsScroll.AutoScroll = $true

    # Apps comuns pre-populados
    $commonApps = @(
        @{ Exe="chrome.exe";          Name="Google Chrome";           Cat="Navegadores" }
        @{ Exe="msedge.exe";          Name="Microsoft Edge";          Cat="Navegadores" }
        @{ Exe="firefox.exe";         Name="Mozilla Firefox";         Cat="Navegadores" }
        @{ Exe="opera.exe";           Name="Opera";                   Cat="Navegadores" }
        @{ Exe="brave.exe";           Name="Brave Browser";           Cat="Navegadores" }

        @{ Exe="cmd.exe";             Name="Prompt de Comando";       Cat="Sistema" }
        @{ Exe="powershell.exe";      Name="PowerShell";              Cat="Sistema" }
        @{ Exe="regedit.exe";         Name="Editor de Registro";      Cat="Sistema" }
        @{ Exe="taskmgr.exe";         Name="Gerenciador de Tarefas";  Cat="Sistema" }
        @{ Exe="mmc.exe";             Name="Console de Gerenciamento";Cat="Sistema" }
        @{ Exe="control.exe";         Name="Painel de Controle";      Cat="Sistema" }
        @{ Exe="msconfig.exe";        Name="Configuracao do Sistema"; Cat="Sistema" }
        @{ Exe="gpedit.msc";          Name="Editor de GPO Local";     Cat="Sistema" }

        @{ Exe="mstsc.exe";           Name="Area de Trabalho Remota"; Cat="Rede/Remoto" }
        @{ Exe="putty.exe";           Name="PuTTY (SSH)";             Cat="Rede/Remoto" }
        @{ Exe="teamviewer.exe";      Name="TeamViewer";              Cat="Rede/Remoto" }
        @{ Exe="anydesk.exe";         Name="AnyDesk";                 Cat="Rede/Remoto" }

        @{ Exe="notepad.exe";         Name="Bloco de Notas";          Cat="Aplicativos" }
        @{ Exe="wordpad.exe";         Name="WordPad";                 Cat="Aplicativos" }
        @{ Exe="calc.exe";            Name="Calculadora";             Cat="Aplicativos" }
        @{ Exe="mspaint.exe";         Name="Paint";                   Cat="Aplicativos" }
        @{ Exe="snippingtool.exe";    Name="Ferramenta de Recorte";   Cat="Aplicativos" }

        @{ Exe="telegram.exe";        Name="Telegram Desktop";        Cat="Comunicacao" }
        @{ Exe="whatsapp.exe";        Name="WhatsApp Desktop";        Cat="Comunicacao" }
        @{ Exe="discord.exe";         Name="Discord";                 Cat="Comunicacao" }
        @{ Exe="slack.exe";           Name="Slack";                   Cat="Comunicacao" }
        @{ Exe="teams.exe";           Name="Microsoft Teams (antigo)";Cat="Comunicacao" }
        @{ Exe="ms-teams.exe";        Name="Microsoft Teams (novo)";  Cat="Comunicacao" }

        @{ Exe="spotify.exe";         Name="Spotify";                 Cat="Midia/Jogos" }
        @{ Exe="steam.exe";           Name="Steam";                   Cat="Midia/Jogos" }
        @{ Exe="epicgameslauncher.exe";Name="Epic Games";             Cat="Midia/Jogos" }
        @{ Exe="vlc.exe";             Name="VLC Media Player";        Cat="Midia/Jogos" }

        @{ Exe="torrent.exe";         Name="Cliente Torrent";         Cat="Downloads" }
        @{ Exe="qbittorrent.exe";     Name="qBittorrent";             Cat="Downloads" }
        @{ Exe="utorrent.exe";        Name="uTorrent";                Cat="Downloads" }
    )

    # Carregar apps ja bloqueados do SYSVOL
    $script:blockedApps = [System.Collections.ArrayList]@()
    $blockFile = "$sysvolPath\Machine\blocked_apps.txt"
    if ($sysvolPath -and (Test-Path $blockFile)) {
        Get-Content $blockFile | ForEach-Object { if ($_.Trim()) { $script:blockedApps.Add($_.Trim().ToLower()) | Out-Null } }
    }

    $ay = 10
    $lastAppCat = ""
    $appChecks = @{}
    foreach ($app in $commonApps) {
        if ($app.Cat -ne $lastAppCat) {
            if ($lastAppCat -ne "") { $ay += 6 }
            $acl = New-Object System.Windows.Forms.Label; $acl.Text = $app.Cat.ToUpper(); $acl.AutoSize = $true
            $acl.ForeColor = $Accent; $acl.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
            $acl.Location = New-Object System.Drawing.Point(15, $ay); $pnlAppsScroll.Controls.Add($acl); $ay += 22
            $lastAppCat = $app.Cat
        }
        $achk = New-Object System.Windows.Forms.CheckBox
        $achk.Text = "  $($app.Name)  ($($app.Exe))"
        $achk.AutoSize = $true; $achk.ForeColor = $FgText
        $achk.Font = New-Object System.Drawing.Font("Segoe UI", 10)
        $achk.Location = New-Object System.Drawing.Point(25, $ay)
        $achk.Tag = $app.Exe
        if ($script:blockedApps -contains $app.Exe.ToLower()) { $achk.Checked = $true }
        $pnlAppsScroll.Controls.Add($achk); $ay += 28
        $appChecks[$app.Exe] = $achk
    }

    # Secao personalizada
    $ay += 10
    $customLbl = New-Object System.Windows.Forms.Label; $customLbl.Text = "PERSONALIZADO"; $customLbl.AutoSize = $true
    $customLbl.ForeColor = $Yellow; $customLbl.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $customLbl.Location = New-Object System.Drawing.Point(15, $ay); $pnlAppsScroll.Controls.Add($customLbl); $ay += 24

    # ── Apps do Cadastro (apps_db.json) ──
    $catalogApps = @()
    if (Test-Path $script:AppsDbPath) {
        try {
            $catalogApps = @(Get-Content $script:AppsDbPath -Raw -Encoding UTF8 | ConvertFrom-Json)
        } catch {}
    }
    if ($catalogApps.Count -gt 0) {
        $catLblDB = New-Object System.Windows.Forms.Label; $catLblDB.Text = "DO CADASTRO DE APPS ($($catalogApps.Count) apps cadastrados)"
        $catLblDB.AutoSize = $true; $catLblDB.ForeColor = [System.Drawing.Color]::FromArgb(203,166,247)
        $catLblDB.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $catLblDB.Location = New-Object System.Drawing.Point(15, $ay); $pnlAppsScroll.Controls.Add($catLblDB); $ay += 24
        foreach ($cApp in $catalogApps) {
            $exeName = ""
            # Tentar extrair exe do link/caminho
            if ($cApp.Link) {
                $exeName = [System.IO.Path]::GetFileName($cApp.Link)
                if (-not $exeName -or -not $exeName.EndsWith(".exe")) { $exeName = "$($cApp.Name -replace '\s','').exe" }
            } else {
                $exeName = "$($cApp.Name -replace '\s','').exe"
            }
            $exeName = $exeName.ToLower()
            # Pular se ja esta na lista comum
            if ($knownExes -contains $exeName) { continue }
            $catChk = New-Object System.Windows.Forms.CheckBox
            $catChk.Text = "  $($cApp.Name)  ($exeName)  [$($cApp.Category)]"
            $catChk.AutoSize = $true; $catChk.ForeColor = $FgText
            $catChk.Font = New-Object System.Drawing.Font("Segoe UI", 10)
            $catChk.Location = New-Object System.Drawing.Point(25, $ay)
            $catChk.Tag = $exeName
            if ($script:blockedApps -contains $exeName) { $catChk.Checked = $true }
            $pnlAppsScroll.Controls.Add($catChk); $ay += 28
            $appChecks[$exeName] = $catChk
        }
        $ay += 10
    }

    $custLbl2 = New-Object System.Windows.Forms.Label; $custLbl2.Text = "Adicione executaveis extras (um por linha):"
    $custLbl2.AutoSize = $true; $custLbl2.ForeColor = [System.Drawing.Color]::FromArgb(108,112,134)
    $custLbl2.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $custLbl2.Location = New-Object System.Drawing.Point(25, $ay); $pnlAppsScroll.Controls.Add($custLbl2); $ay += 22

    $txtCustomApps = New-Object System.Windows.Forms.TextBox
    $txtCustomApps.Multiline = $true; $txtCustomApps.ScrollBars = "Vertical"
    $txtCustomApps.Location = New-Object System.Drawing.Point(25, $ay)
    $txtCustomApps.Size = New-Object System.Drawing.Size(740, 80)
    $txtCustomApps.BackColor = $BgField; $txtCustomApps.ForeColor = $FgText
    $txtCustomApps.Font = New-Object System.Drawing.Font("Consolas", 10)
    # Pre-popular com apps bloqueados que nao estao na lista comum
    $knownExes = $commonApps | ForEach-Object { $_.Exe.ToLower() }
    $customExes = $script:blockedApps | Where-Object { $_ -notin $knownExes }
    if ($customExes) { $txtCustomApps.Text = ($customExes -join "`r`n") }
    $pnlAppsScroll.Controls.Add($txtCustomApps)

    $pnlAppsMain.Controls.Add($pnlAppsScroll)
    $tabApps.Controls.Add($pnlAppsMain)
    $tabs.TabPages.Add($tabApps)

    # ════════════════════════════════════════
    #  TAB 5: SCRIPTS E INSTALACAO
    # ════════════════════════════════════════
    $tabReg = New-Object System.Windows.Forms.TabPage
    $tabReg.Text = "  Scripts / Instalacao  "
    $tabReg.BackColor = $BgDark

    $pnlTab5 = New-Object System.Windows.Forms.Panel
    $pnlTab5.Dock = "Fill"; $pnlTab5.AutoScroll = $true; $pnlTab5.Padding = New-Object System.Windows.Forms.Padding(15)

    # ── SECAO 1: INSTALACAO DE SOFTWARE ──
    $lblInstTitle = New-Object System.Windows.Forms.Label
    $lblInstTitle.Text = "INSTALACAO DE SOFTWARE"
    $lblInstTitle.AutoSize = $true; $lblInstTitle.ForeColor = $Yellow
    $lblInstTitle.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $lblInstTitle.Location = New-Object System.Drawing.Point(15, 10)
    $pnlTab5.Controls.Add($lblInstTitle)

    $lblInstDesc = New-Object System.Windows.Forms.Label
    $lblInstDesc.Text = "Adicione links ou caminhos de instaladores. Cole o caminho de rede ou URL do executavel."
    $lblInstDesc.AutoSize = $true; $lblInstDesc.ForeColor = $Overlay
    $lblInstDesc.Location = New-Object System.Drawing.Point(15, 36)
    $pnlTab5.Controls.Add($lblInstDesc)

    $toolInst = New-Object System.Windows.Forms.Panel
    $toolInst.Location = New-Object System.Drawing.Point(10, 58); $toolInst.Size = New-Object System.Drawing.Size(950, 38)

    $btnAddInst = New-Object System.Windows.Forms.Button
    $btnAddInst.Text = "+ Adicionar Software"; $btnAddInst.Location = New-Object System.Drawing.Point(5, 4)
    $btnAddInst.Size = New-Object System.Drawing.Size(180, 30); $btnAddInst.BackColor = $Green; $btnAddInst.ForeColor = $BgDark
    $btnAddInst.FlatStyle = "Flat"; $btnAddInst.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnAddInst.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $toolInst.Controls.Add($btnAddInst)

    $btnRemoveInst = New-Object System.Windows.Forms.Button
    $btnRemoveInst.Text = "Remover"; $btnRemoveInst.Location = New-Object System.Drawing.Point(195, 4)
    $btnRemoveInst.Size = New-Object System.Drawing.Size(100, 30); $btnRemoveInst.BackColor = $Red; $btnRemoveInst.ForeColor = [System.Drawing.Color]::White
    $btnRemoveInst.FlatStyle = "Flat"; $btnRemoveInst.Cursor = [System.Windows.Forms.Cursors]::Hand
    $toolInst.Controls.Add($btnRemoveInst)
    $pnlTab5.Controls.Add($toolInst)

    $dgvInst = New-Object System.Windows.Forms.DataGridView
    $dgvInst.Location = New-Object System.Drawing.Point(15, 100); $dgvInst.Size = New-Object System.Drawing.Size(940, 200)
    $dgvInst.BackgroundColor = $BgDark; $dgvInst.ForeColor = $FgText; $dgvInst.GridColor = $BgField
    $dgvInst.BorderStyle = "FixedSingle"; $dgvInst.CellBorderStyle = "SingleHorizontal"
    $dgvInst.ColumnHeadersDefaultCellStyle.BackColor = $BgPanel; $dgvInst.ColumnHeadersDefaultCellStyle.ForeColor = $Accent
    $dgvInst.DefaultCellStyle.BackColor = $BgDark; $dgvInst.DefaultCellStyle.ForeColor = $FgText
    $dgvInst.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(60,60,90)
    $dgvInst.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(36, 36, 54)
    $dgvInst.RowHeadersVisible = $false; $dgvInst.AllowUserToAddRows = $false
    $dgvInst.SelectionMode = "FullRowSelect"; $dgvInst.AutoSizeColumnsMode = "Fill"
    $dgvInst.EnableHeadersVisualStyles = $false; $dgvInst.ReadOnly = $true
    $dgvInst.Anchor = "Top,Left,Right"

    $dgvInst.Columns.Add("InstName", "Nome") | Out-Null
    $dgvInst.Columns.Add("InstPath", "Caminho / URL") | Out-Null
    $dgvInst.Columns.Add("InstArgs", "Argumentos") | Out-Null
    $dgvInst.Columns["InstName"].FillWeight = 20; $dgvInst.Columns["InstPath"].FillWeight = 55; $dgvInst.Columns["InstArgs"].FillWeight = 25

    $script:installEntries = [System.Collections.ArrayList]@()
    # Carregar instalacoes salvas (nosso JSON)
    $instFile = "$sysvolPath\Machine\install_scripts.json"
    if ($sysvolPath -and (Test-Path $instFile)) {
        try {
            $loaded = Get-Content $instFile -Raw | ConvertFrom-Json
            foreach ($r in $loaded) {
                $entry = @{ Name=$r.Name; Path=$r.Path; Args=$r.Args }
                $script:installEntries.Add($entry) | Out-Null
                $dgvInst.Rows.Add($r.Name, $r.Path, $r.Args) | Out-Null
            }
        } catch {}
    }

    # Carregar scripts nativos detectados (de $script:nativeScripts)
    foreach ($ns in $script:nativeScripts) {
        $alreadyListed = $false
        foreach ($e in $script:installEntries) {
            if ($e.Path -and $e.Path -like "*$($ns.Command)*") { $alreadyListed = $true; break }
        }
        if (-not $alreadyListed -and $ns.Command) {
            $entry = @{ Name=[System.IO.Path]::GetFileNameWithoutExtension($ns.Command); Path=$ns.Command; Args="($($ns.Type)) $($ns.Parameters)".Trim() }
            $script:installEntries.Add($entry) | Out-Null
            $rowIdx = $dgvInst.Rows.Add($entry.Name, $entry.Path, $entry.Args)
            $dgvInst.Rows[$rowIdx].DefaultCellStyle.ForeColor = $Yellow
        }
    }

    # Detectar scripts de startup existentes na pasta SYSVOL (configurados via GPMC)
    if ($sysvolPath) {
        $startupDir = "$sysvolPath\Machine\Scripts\Startup"
        if (Test-Path $startupDir) {
            $scripts = Get-ChildItem $startupDir -File -ErrorAction SilentlyContinue
            foreach ($sf in $scripts) {
                if ($sf.Name -eq "scripts.ini" -or $sf.Name -eq "psscripts.ini") { continue }
                $alreadyListed = $false
                foreach ($e in $script:installEntries) {
                    if ($e.Path -and $e.Path -like "*$($sf.Name)*") { $alreadyListed = $true; break }
                }
                if (-not $alreadyListed) {
                    $entry = @{ Name=$sf.BaseName; Path=$sf.FullName; Args="(startup script SYSVOL)" }
                    $script:installEntries.Add($entry) | Out-Null
                    $dgvInst.Rows.Add($entry.Name, $entry.Path, $entry.Args) | Out-Null
                }
            }
        }

        # Ler scripts.ini para pegar scripts registrados com parametros
        foreach ($iniSub in @("Machine\Scripts\Startup\scripts.ini","Machine\Scripts\scripts.ini","MACHINE\Scripts\Startup\scripts.ini","MACHINE\Scripts\scripts.ini")) {
            $scriptsIni = "$sysvolPath\$iniSub"
            if (Test-Path $scriptsIni) {
                try {
                    $iniContent = Get-Content $scriptsIni -Encoding Unicode -ErrorAction SilentlyContinue
                    if (-not $iniContent) { $iniContent = Get-Content $scriptsIni -ErrorAction SilentlyContinue }
                    $idx = 0
                    while ($true) {
                        $cmdLine = $iniContent | Where-Object { $_ -match "^${idx}CmdLine=" }
                        $parLine = $iniContent | Where-Object { $_ -match "^${idx}Parameters=" }
                        if (-not $cmdLine) { break }
                        $cmd = ($cmdLine -split "=", 2)[1].Trim()
                        $par = if ($parLine) { ($parLine -split "=", 2)[1].Trim() } else { "" }
                        if ($cmd) {
                            $alreadyListed = $false
                            foreach ($e in $script:installEntries) {
                                if ($e.Path -and ($e.Path -like "*$cmd*" -or $cmd -like "*$($e.Name)*")) { $alreadyListed = $true; break }
                            }
                            if (-not $alreadyListed) {
                                $entry = @{ Name=[System.IO.Path]::GetFileNameWithoutExtension($cmd); Path=$cmd; Args=$par }
                                $script:installEntries.Add($entry) | Out-Null
                                $dgvInst.Rows.Add($entry.Name, $entry.Path, $entry.Args) | Out-Null
                            }
                        }
                        $idx++
                    }
                } catch {}
                break
            }
        }

        # Detectar pacotes MSI de Software Installation (via AD Class Store)
        try {
            $pkgPath = "LDAP://CN=Packages,CN=Class Store,CN=Machine,$($gpoEntry.distinguishedName)"
            $pkgContainer = [ADSI]$pkgPath
            if ($pkgContainer -and $pkgContainer.Children) {
                foreach ($pkg in $pkgContainer.Children) {
                    $pkgName = [string]$pkg.displayName
                    $msiPath = [string]$pkg.msiFileList
                    if (-not $msiPath) { $msiPath = [string]$pkg.msiScriptPath }
                    if ($pkgName) {
                        $alreadyListed = $false
                        foreach ($e in $script:installEntries) {
                            if ($e.Name -eq $pkgName) { $alreadyListed = $true; break }
                        }
                        if (-not $alreadyListed) {
                            $entry = @{ Name=$pkgName; Path=$(if($msiPath){$msiPath}else{"(MSI via GPO)"}); Args="(Software Installation)" }
                            $script:installEntries.Add($entry) | Out-Null
                            $dgvInst.Rows.Add($entry.Name, $entry.Path, $entry.Args) | Out-Null
                        }
                    }
                }
            }
        } catch {}
    }

    $pnlTab5.Controls.Add($dgvInst)

    # ── SECAO 2: REGISTRO AVANCADO ──
    $lblRegTitle = New-Object System.Windows.Forms.Label
    $lblRegTitle.Text = "REGISTRO (nativo + avancado)"
    $lblRegTitle.AutoSize = $true; $lblRegTitle.ForeColor = $Overlay
    $lblRegTitle.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $lblRegTitle.Location = New-Object System.Drawing.Point(15, 315)
    $pnlTab5.Controls.Add($lblRegTitle)

    $toolReg = New-Object System.Windows.Forms.Panel
    $toolReg.Location = New-Object System.Drawing.Point(10, 340); $toolReg.Size = New-Object System.Drawing.Size(950, 38)

    $btnAddReg = New-Object System.Windows.Forms.Button
    $btnAddReg.Text = "+ Regra de Registro"; $btnAddReg.Location = New-Object System.Drawing.Point(5, 4)
    $btnAddReg.Size = New-Object System.Drawing.Size(160, 30); $btnAddReg.BackColor = $BgField; $btnAddReg.ForeColor = $FgText
    $btnAddReg.FlatStyle = "Flat"; $btnAddReg.Cursor = [System.Windows.Forms.Cursors]::Hand
    $toolReg.Controls.Add($btnAddReg)

    $btnRemoveReg = New-Object System.Windows.Forms.Button
    $btnRemoveReg.Text = "Remover"; $btnRemoveReg.Location = New-Object System.Drawing.Point(175, 4)
    $btnRemoveReg.Size = New-Object System.Drawing.Size(100, 30); $btnRemoveReg.BackColor = $Red; $btnRemoveReg.ForeColor = [System.Drawing.Color]::White
    $btnRemoveReg.FlatStyle = "Flat"; $btnRemoveReg.Cursor = [System.Windows.Forms.Cursors]::Hand
    $toolReg.Controls.Add($btnRemoveReg)
    $pnlTab5.Controls.Add($toolReg)

    $dgvReg = New-Object System.Windows.Forms.DataGridView
    $dgvReg.Location = New-Object System.Drawing.Point(15, 382); $dgvReg.Size = New-Object System.Drawing.Size(940, 180)
    $dgvReg.BackgroundColor = $BgDark; $dgvReg.ForeColor = $FgText; $dgvReg.GridColor = $BgField
    $dgvReg.BorderStyle = "FixedSingle"; $dgvReg.CellBorderStyle = "SingleHorizontal"
    $dgvReg.ColumnHeadersDefaultCellStyle.BackColor = $BgPanel; $dgvReg.ColumnHeadersDefaultCellStyle.ForeColor = $Accent
    $dgvReg.DefaultCellStyle.BackColor = $BgDark; $dgvReg.DefaultCellStyle.ForeColor = $FgText
    $dgvReg.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(60,60,90)
    $dgvReg.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(36, 36, 54)
    $dgvReg.RowHeadersVisible = $false; $dgvReg.AllowUserToAddRows = $false
    $dgvReg.SelectionMode = "FullRowSelect"; $dgvReg.AutoSizeColumnsMode = "Fill"
    $dgvReg.EnableHeadersVisualStyles = $false; $dgvReg.ReadOnly = $true
    $dgvReg.Anchor = "Top,Left,Right"

    $dgvReg.Columns.Add("Hive", "Hive") | Out-Null
    $dgvReg.Columns.Add("Key", "Caminho") | Out-Null
    $dgvReg.Columns.Add("ValueName", "Valor") | Out-Null
    $dgvReg.Columns.Add("ValueData", "Dado") | Out-Null
    $dgvReg.Columns.Add("Type", "Tipo") | Out-Null
    $dgvReg.Columns["Hive"].FillWeight = 8; $dgvReg.Columns["Key"].FillWeight = 38
    $dgvReg.Columns["ValueName"].FillWeight = 20; $dgvReg.Columns["ValueData"].FillWeight = 24; $dgvReg.Columns["Type"].FillWeight = 10

    $script:regEntries = [System.Collections.ArrayList]@()
    $regFile = "$sysvolPath\Machine\registry_rules.json"
    if ($sysvolPath -and (Test-Path $regFile)) {
        try {
            $loaded = Get-Content $regFile -Raw | ConvertFrom-Json
            foreach ($r in $loaded) {
                $entry = @{ Hive=$r.Hive; Key=$r.Key; ValueName=$r.ValueName; ValueData=$r.ValueData; Type=$r.Type }
                $script:regEntries.Add($entry) | Out-Null
                $dgvReg.Rows.Add($r.Hive, $r.Key, $r.ValueName, $r.ValueData, $r.Type) | Out-Null
            }
        } catch {}
    }

    # Carregar entradas nativas do Registry.pol que NAO foram mapeadas
    foreach ($nr in $script:nativeRegEntries) {
        $alreadyListed = $false
        foreach ($e in $script:regEntries) {
            if ($e.Key -eq $nr.Key -and $e.ValueName -eq $nr.ValueName) { $alreadyListed = $true; break }
        }
        # Tambem ignorar se ja foi mapeado para Politicas Comuns
        $isMapped = $false
        foreach ($pol in $commonPolicies) {
            if ($nr.Key -like "*$($pol.Key)*" -and $nr.ValueName -eq $pol.Val) { $isMapped = $true; break }
        }
        if (-not $alreadyListed -and -not $isMapped) {
            $entry = @{ Hive=$nr.Hive; Key=$nr.Key; ValueName=$nr.ValueName; ValueData=$nr.ValueData; Type=$nr.Type }
            $script:regEntries.Add($entry) | Out-Null
            $rowIdx = $dgvReg.Rows.Add($nr.Hive, $nr.Key, $nr.ValueName, $nr.ValueData, $nr.Type)
            $dgvReg.Rows[$rowIdx].DefaultCellStyle.ForeColor = $Yellow
        }
    }

    $pnlTab5.Controls.Add($dgvReg)

    $tabReg.Controls.Add($pnlTab5)
    $tabs.TabPages.Add($tabReg)

    # ════════════════════════════════════════
    #  TAB 6: PREFERENCIAS (GPP)
    # ════════════════════════════════════════
    $tabPref = New-Object System.Windows.Forms.TabPage
    $tabPref.Text = "  Preferencias (GPP)  "
    $tabPref.BackColor = $BgDark

    # Painel header com labels (altura fixa, nao Dock)
    $pnlPrefHeader = New-Object System.Windows.Forms.Panel
    $pnlPrefHeader.Dock = "Top"; $pnlPrefHeader.Height = 58; $pnlPrefHeader.BackColor = $BgDark

    $lblPrefTitle = New-Object System.Windows.Forms.Label
    $lblPrefTitle.Text = "PREFERENCIAS NATIVAS (GROUP POLICY PREFERENCES)"
    $lblPrefTitle.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $lblPrefTitle.ForeColor = $Accent; $lblPrefTitle.AutoSize = $true
    $lblPrefTitle.Location = New-Object System.Drawing.Point(15, 8)
    $pnlPrefHeader.Controls.Add($lblPrefTitle)

    $lblPrefDesc = New-Object System.Windows.Forms.Label
    $lblPrefDesc.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblPrefDesc.AutoSize = $true
    $lblPrefDesc.Location = New-Object System.Drawing.Point(15, 34)

    $prefCount = $script:nativePreferences.Count
    if ($prefCount -gt 0) {
        $tipos = ($script:nativePreferences | ForEach-Object { $_.Tipo } | Sort-Object -Unique) -join ", "
        $lblPrefDesc.Text = "$prefCount configuracao(oes) encontrada(s): $tipos"
        $lblPrefDesc.ForeColor = $Green
    } else {
        $lblPrefDesc.Text = "Nenhuma preferencia nativa encontrada nesta GPO. (Atalhos, Drives, Impressoras, Arquivos, etc)"
        $lblPrefDesc.ForeColor = $Overlay
    }
    $pnlPrefHeader.Controls.Add($lblPrefDesc)

    $dgvPref = New-Object System.Windows.Forms.DataGridView
    $dgvPref.Dock = "Fill"
    $dgvPref.BackgroundColor = $BgDark; $dgvPref.GridColor = $Overlay; $dgvPref.ForeColor = $FgText
    $dgvPref.DefaultCellStyle.BackColor = $BgField; $dgvPref.DefaultCellStyle.ForeColor = $FgText
    $dgvPref.DefaultCellStyle.SelectionBackColor = $Accent; $dgvPref.DefaultCellStyle.SelectionForeColor = $BgDark
    $dgvPref.ColumnHeadersDefaultCellStyle.BackColor = $BgPanel; $dgvPref.ColumnHeadersDefaultCellStyle.ForeColor = $Accent
    $dgvPref.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $dgvPref.EnableHeadersVisualStyles = $false
    $dgvPref.RowHeadersVisible = $false; $dgvPref.AllowUserToAddRows = $false
    $dgvPref.ReadOnly = $true; $dgvPref.SelectionMode = "FullRowSelect"
    $dgvPref.AutoSizeColumnsMode = "None"; $dgvPref.RowTemplate.Height = 28
    $dgvPref.BorderStyle = "None"; $dgvPref.CellBorderStyle = "SingleHorizontal"

    $dgvPref.Columns.Add("Tipo", "Tipo") | Out-Null
    $dgvPref.Columns.Add("Escopo", "Escopo") | Out-Null
    $dgvPref.Columns.Add("Acao", "Acao") | Out-Null
    $dgvPref.Columns.Add("Nome", "Nome") | Out-Null
    $dgvPref.Columns.Add("Caminho", "Caminho / Valor") | Out-Null
    $dgvPref.Columns.Add("Detalhes", "Detalhes") | Out-Null

    $dgvPref.Columns["Tipo"].Width = 120; $dgvPref.Columns["Tipo"].MinimumWidth = 100
    $dgvPref.Columns["Escopo"].Width = 90; $dgvPref.Columns["Escopo"].MinimumWidth = 80
    $dgvPref.Columns["Acao"].Width = 95; $dgvPref.Columns["Acao"].MinimumWidth = 80
    $dgvPref.Columns["Nome"].Width = 200; $dgvPref.Columns["Nome"].AutoSizeMode = "Fill"; $dgvPref.Columns["Nome"].FillWeight = 25; $dgvPref.Columns["Nome"].MinimumWidth = 120
    $dgvPref.Columns["Caminho"].AutoSizeMode = "Fill"; $dgvPref.Columns["Caminho"].FillWeight = 45; $dgvPref.Columns["Caminho"].MinimumWidth = 200
    $dgvPref.Columns["Detalhes"].AutoSizeMode = "Fill"; $dgvPref.Columns["Detalhes"].FillWeight = 30; $dgvPref.Columns["Detalhes"].MinimumWidth = 150

    # Cores por tipo de acao
    $actColors = @{
        "Criar"       = $Green
        "Atualizar"   = $Accent
        "Substituir"  = $Yellow
        "Deletar"     = $Red
    }
    $typeColors = @{ "Atalho"=$Mauve; "Drive Mapeado"=$Green; "Impressora"=$Accent; "Arquivo"=$Yellow; "Pasta"=$Yellow; "Servico"=$Red; "Grupo/Usuario"=$Red; "Registro (Pref)"=$Yellow; "Energia"=$Overlay; "Redir. Pasta"=$Green; "Seguranca"=$Red; "Privilegio"=$Mauve; "Membro Grupo"=$Red; "Reg. Seguranca"=$Yellow; "Perm. Registro"=$Yellow; "Perm. Arquivo"=$Yellow; "Software (MSI)"=$Green }

    foreach ($pref in $script:nativePreferences) {
        $rowIdx = $dgvPref.Rows.Add($pref.Tipo, $pref.Escopo, $pref.Acao, $pref.Nome, $pref.Caminho, $pref.Detalhes)
        if ($actColors[$pref.Acao]) {
            $dgvPref.Rows[$rowIdx].Cells["Acao"].Style.ForeColor = $actColors[$pref.Acao]
        }
        if ($typeColors[$pref.Tipo]) {
            $dgvPref.Rows[$rowIdx].Cells["Tipo"].Style.ForeColor = $typeColors[$pref.Tipo]
        }
    }

    $tabPref.Controls.Add($dgvPref)
    $tabPref.Controls.Add($pnlPrefHeader)
    $pnlPrefHeader.BringToFront()

    $tabs.TabPages.Add($tabPref)

    # ════════════════════════════════════════
    #  BARRA INFERIOR (Salvar/Cancelar)
    # ════════════════════════════════════════
    $bottomPanel = New-Object System.Windows.Forms.Panel
    $bottomPanel.Dock = "Bottom"; $bottomPanel.Height = 55; $bottomPanel.BackColor = $BgPanel

    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = "  SALVAR TUDO  "; $btnSave.Size = New-Object System.Drawing.Size(180, 38)
    $btnSave.Location = New-Object System.Drawing.Point(15, 8)
    $btnSave.BackColor = $Green; $btnSave.ForeColor = $BgDark; $btnSave.FlatStyle = "Flat"
    $btnSave.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $btnSave.Cursor = [System.Windows.Forms.Cursors]::Hand
    $bottomPanel.Controls.Add($btnSave)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancelar"; $btnCancel.Size = New-Object System.Drawing.Size(100, 38)
    $btnCancel.Location = New-Object System.Drawing.Point(205, 8)
    $btnCancel.BackColor = $BgField; $btnCancel.ForeColor = $FgText; $btnCancel.FlatStyle = "Flat"
    $btnCancel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $btnCancel.Cursor = [System.Windows.Forms.Cursors]::Hand; $btnCancel.DialogResult = "Cancel"
    $bottomPanel.Controls.Add($btnCancel)

    $lblSaveStatus = New-Object System.Windows.Forms.Label
    $lblSaveStatus.AutoSize = $true; $lblSaveStatus.Location = New-Object System.Drawing.Point(320, 18)
    $lblSaveStatus.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $bottomPanel.Controls.Add($lblSaveStatus)

    $dlg.Controls.Add($tabs)
    $dlg.Controls.Add($bottomPanel)
    $dlg.CancelButton = $btnCancel

    # ════════════════════════════════════════
    #  EVENTOS
    # ════════════════════════════════════════

    # Adicionar software/instalador
    $btnAddInst.Add_Click({
        $instDlg = New-Object System.Windows.Forms.Form
        $instDlg.Text = "Adicionar Software"; $instDlg.Size = New-Object System.Drawing.Size(580, 280)
        $instDlg.StartPosition = "CenterParent"; $instDlg.BackColor = $BgPanel; $instDlg.ForeColor = $FgText
        $instDlg.FormBorderStyle = "FixedDialog"; $instDlg.MaximizeBox = $false

        $iy = 15
        $il1 = New-Object System.Windows.Forms.Label; $il1.Text = "Nome do Software:"; $il1.AutoSize = $true
        $il1.Location = New-Object System.Drawing.Point(15, $iy); $il1.ForeColor = $Accent; $instDlg.Controls.Add($il1); $iy += 22
        $tInstName = New-Object System.Windows.Forms.TextBox; $tInstName.Location = New-Object System.Drawing.Point(15, $iy)
        $tInstName.Size = New-Object System.Drawing.Size(530, 25); $tInstName.BackColor = $BgField; $tInstName.ForeColor = $FgText
        $tInstName.Font = New-Object System.Drawing.Font("Segoe UI", 10)
        $instDlg.Controls.Add($tInstName); $iy += 38

        $il2 = New-Object System.Windows.Forms.Label; $il2.Text = "Caminho ou URL do instalador:"; $il2.AutoSize = $true
        $il2.Location = New-Object System.Drawing.Point(15, $iy); $il2.ForeColor = $Accent; $instDlg.Controls.Add($il2); $iy += 22
        $tInstPath = New-Object System.Windows.Forms.TextBox; $tInstPath.Location = New-Object System.Drawing.Point(15, $iy)
        $tInstPath.Size = New-Object System.Drawing.Size(530, 25); $tInstPath.BackColor = $BgField; $tInstPath.ForeColor = $FgText
        $tInstPath.Font = New-Object System.Drawing.Font("Segoe UI", 10)
        $instDlg.Controls.Add($tInstPath); $iy += 38

        $il3 = New-Object System.Windows.Forms.Label; $il3.Text = "Argumentos (opcional):    Ex: /silent /norestart"; $il3.AutoSize = $true
        $il3.Location = New-Object System.Drawing.Point(15, $iy); $il3.ForeColor = $Overlay; $instDlg.Controls.Add($il3); $iy += 22
        $tInstArgs = New-Object System.Windows.Forms.TextBox; $tInstArgs.Location = New-Object System.Drawing.Point(15, $iy)
        $tInstArgs.Size = New-Object System.Drawing.Size(530, 25); $tInstArgs.BackColor = $BgField; $tInstArgs.ForeColor = $FgText
        $tInstArgs.Font = New-Object System.Drawing.Font("Segoe UI", 10)
        $instDlg.Controls.Add($tInstArgs); $iy += 38

        $bAddInst = New-Object System.Windows.Forms.Button; $bAddInst.Text = "Adicionar"; $bAddInst.Size = New-Object System.Drawing.Size(140, 35)
        $bAddInst.Location = New-Object System.Drawing.Point(215, $iy); $bAddInst.BackColor = $Green; $bAddInst.ForeColor = $BgDark
        $bAddInst.FlatStyle = "Flat"; $bAddInst.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        $bAddInst.DialogResult = "OK"; $instDlg.Controls.Add($bAddInst)
        $instDlg.AcceptButton = $bAddInst

        if ($instDlg.ShowDialog() -eq "OK" -and $tInstPath.Text.Trim()) {
            $name = if ($tInstName.Text.Trim()) { $tInstName.Text.Trim() } else { [System.IO.Path]::GetFileName($tInstPath.Text.Trim()) }
            $entry = @{ Name=$name; Path=$tInstPath.Text.Trim(); Args=$tInstArgs.Text.Trim() }
            $script:installEntries.Add($entry) | Out-Null
            $dgvInst.Rows.Add($entry.Name, $entry.Path, $entry.Args) | Out-Null
        }
        $instDlg.Dispose()
    })

    $btnRemoveInst.Add_Click({
        if ($dgvInst.SelectedRows.Count -gt 0) {
            $idx = $dgvInst.SelectedRows[0].Index
            $dgvInst.Rows.RemoveAt($idx)
            if ($idx -lt $script:installEntries.Count) { $script:installEntries.RemoveAt($idx) }
        }
    })

    # Adicionar regra customizada de registro
    $btnAddReg.Add_Click({
        $regDlg = New-Object System.Windows.Forms.Form
        $regDlg.Text = "Nova Regra de Registro"; $regDlg.Size = New-Object System.Drawing.Size(550, 330)
        $regDlg.StartPosition = "CenterParent"; $regDlg.BackColor = $BgPanel; $regDlg.ForeColor = $FgText
        $regDlg.FormBorderStyle = "FixedDialog"; $regDlg.MaximizeBox = $false

        $ry = 15
        $rl1 = New-Object System.Windows.Forms.Label; $rl1.Text = "Hive:"; $rl1.AutoSize = $true
        $rl1.Location = New-Object System.Drawing.Point(15, $ry); $regDlg.Controls.Add($rl1); $ry += 22
        $cmbHive = New-Object System.Windows.Forms.ComboBox
        $cmbHive.Items.AddRange(@("HKLM", "HKCU")); $cmbHive.SelectedIndex = 0; $cmbHive.DropDownStyle = "DropDownList"
        $cmbHive.Location = New-Object System.Drawing.Point(15, $ry); $cmbHive.Size = New-Object System.Drawing.Size(150, 25)
        $cmbHive.BackColor = $BgField; $cmbHive.ForeColor = $FgText; $regDlg.Controls.Add($cmbHive); $ry += 35

        $rl2 = New-Object System.Windows.Forms.Label; $rl2.Text = "Caminho:"; $rl2.AutoSize = $true
        $rl2.Location = New-Object System.Drawing.Point(15, $ry); $regDlg.Controls.Add($rl2); $ry += 22
        $tKey = New-Object System.Windows.Forms.TextBox; $tKey.Location = New-Object System.Drawing.Point(15, $ry)
        $tKey.Size = New-Object System.Drawing.Size(500, 25); $tKey.BackColor = $BgField; $tKey.ForeColor = $FgText
        $regDlg.Controls.Add($tKey); $ry += 35

        $rl3 = New-Object System.Windows.Forms.Label; $rl3.Text = "Nome do Valor:"; $rl3.AutoSize = $true
        $rl3.Location = New-Object System.Drawing.Point(15, $ry); $regDlg.Controls.Add($rl3)
        $rl4 = New-Object System.Windows.Forms.Label; $rl4.Text = "Dado:"; $rl4.AutoSize = $true
        $rl4.Location = New-Object System.Drawing.Point(270, $ry); $regDlg.Controls.Add($rl4); $ry += 22
        $tValName = New-Object System.Windows.Forms.TextBox; $tValName.Location = New-Object System.Drawing.Point(15, $ry)
        $tValName.Size = New-Object System.Drawing.Size(240, 25); $tValName.BackColor = $BgField; $tValName.ForeColor = $FgText
        $regDlg.Controls.Add($tValName)
        $tValData = New-Object System.Windows.Forms.TextBox; $tValData.Location = New-Object System.Drawing.Point(270, $ry)
        $tValData.Size = New-Object System.Drawing.Size(245, 25); $tValData.BackColor = $BgField; $tValData.ForeColor = $FgText
        $regDlg.Controls.Add($tValData); $ry += 35

        $rl5 = New-Object System.Windows.Forms.Label; $rl5.Text = "Tipo:"; $rl5.AutoSize = $true
        $rl5.Location = New-Object System.Drawing.Point(15, $ry); $regDlg.Controls.Add($rl5); $ry += 22
        $cmbType = New-Object System.Windows.Forms.ComboBox
        $cmbType.Items.AddRange(@("String", "DWord", "ExpandString", "MultiString", "QWord")); $cmbType.SelectedIndex = 1
        $cmbType.DropDownStyle = "DropDownList"; $cmbType.Location = New-Object System.Drawing.Point(15, $ry)
        $cmbType.Size = New-Object System.Drawing.Size(150, 25); $cmbType.BackColor = $BgField; $cmbType.ForeColor = $FgText
        $regDlg.Controls.Add($cmbType)

        $bAdd = New-Object System.Windows.Forms.Button; $bAdd.Text = "Adicionar"; $bAdd.Size = New-Object System.Drawing.Size(120, 32)
        $bAdd.Location = New-Object System.Drawing.Point(200, $ry); $bAdd.BackColor = $Green; $bAdd.ForeColor = $BgDark
        $bAdd.FlatStyle = "Flat"; $bAdd.DialogResult = "OK"; $regDlg.Controls.Add($bAdd)
        $regDlg.AcceptButton = $bAdd

        if ($regDlg.ShowDialog() -eq "OK" -and $tKey.Text.Trim()) {
            $entry = @{ Hive=$cmbHive.Text; Key=$tKey.Text.Trim(); ValueName=$tValName.Text.Trim(); ValueData=$tValData.Text; Type=$cmbType.Text }
            $script:regEntries.Add($entry) | Out-Null
            $dgvReg.Rows.Add($entry.Hive, $entry.Key, $entry.ValueName, $entry.ValueData, $entry.Type) | Out-Null
        }
        $regDlg.Dispose()
    })

    $btnRemoveReg.Add_Click({
        if ($dgvReg.SelectedRows.Count -gt 0) {
            $idx = $dgvReg.SelectedRows[0].Index
            $dgvReg.Rows.RemoveAt($idx)
            if ($idx -lt $script:regEntries.Count) { $script:regEntries.RemoveAt($idx) }
        }
    })

    # ════════════════════════════════════════
    #  SALVAR TUDO
    # ════════════════════════════════════════
    $btnSave.Add_Click({
        try {
            $dlg.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            $lblSaveStatus.ForeColor = $Yellow; $lblSaveStatus.Text = "Salvando..."
            [System.Windows.Forms.Application]::DoEvents()
            $changes = @()

            # 1) Salvar nome
            $newName = $tName.Text.Trim()
            if ($newName -and $newName -ne $GpoName) {
                $gpoEntry.Put("displayName", $newName); $changes += "Nome"
            }

            # 1b) Salvar descricao
            $newDesc = $tDesc.Text.Trim()
            if ($newDesc -ne $curDesc) {
                if ($newDesc) { $gpoEntry.Put("description", $newDesc) }
                else { try { $gpoEntry.PutEx(1, "description", $null) } catch {} }
                $changes += "Descricao"
            }

            # 2) Salvar status
            $newFlags = $cmbStatus.SelectedIndex
            if ($newFlags -ne $curFlags) {
                $gpoEntry.Put("flags", $newFlags); $changes += "Status"
            }

            if ($changes.Count -gt 0) { $gpoEntry.SetInfo() }

            # 3) Garantir pasta SYSVOL
            if ($sysvolPath) {
                $machPath = "$sysvolPath\Machine"
                if (-not (Test-Path $machPath)) { New-Item -Path $machPath -ItemType Directory -Force | Out-Null }

                # 4) Salvar politicas comuns como config
                $polConfig = @{}
                foreach ($polDef in $commonPolicies) {
                    $c = $chkBoxes[$polDef.Id]
                    if ($c -and $c.Checked) { $polConfig[$polDef.Id] = $true }
                }
                $polConfig | ConvertTo-Json -Depth 3 | Out-File -FilePath "$machPath\gpo_config.json" -Encoding UTF8
                $enabledCount = ($polConfig.Keys | Measure-Object).Count
                if ($enabledCount -gt 0) { $changes += "$enabledCount politicas" }

                # 5) Salvar apps bloqueados
                $allBlocked = [System.Collections.ArrayList]@()
                foreach ($ak in $appChecks.Keys) {
                    if ($appChecks[$ak].Checked) { $allBlocked.Add($ak.ToLower()) | Out-Null }
                }
                # Adicionar custom
                $txtCustomApps.Text.Split("`n") | ForEach-Object {
                    $x = $_.Trim()
                    if ($x -and $x -notin $allBlocked) { $allBlocked.Add($x.ToLower()) | Out-Null }
                }
                $blockFilePath = "$machPath\blocked_apps.txt"
                if ($allBlocked.Count -gt 0) {
                    $allBlocked | Out-File -FilePath $blockFilePath -Encoding UTF8
                    $changes += "$($allBlocked.Count) apps bloqueados"
                } elseif (Test-Path $blockFilePath) {
                    Remove-Item $blockFilePath -Force; $changes += "Bloqueio limpo"
                }

                # 6) Salvar registry custom
                if ($script:regEntries.Count -gt 0) {
                    $regFilePath = "$machPath\registry_rules.json"
                    $script:regEntries | ConvertTo-Json -Depth 5 | Out-File -FilePath $regFilePath -Encoding UTF8
                    $changes += "$($script:regEntries.Count) regras registro"
                }

                # 7) Salvar instalacoes de software
                if ($script:installEntries.Count -gt 0) {
                    $instFilePath = "$machPath\install_scripts.json"
                    $script:installEntries | ConvertTo-Json -Depth 5 | Out-File -FilePath $instFilePath -Encoding UTF8
                    $changes += "$($script:installEntries.Count) instalacoes"
                }
            }

            if ($changes.Count -gt 0) {
                $lblSaveStatus.ForeColor = $Green
                $lblSaveStatus.Text = "Salvo! (" + ($changes -join ", ") + ")"
                Load-GPOs
            } else {
                $lblSaveStatus.ForeColor = $Yellow
                $lblSaveStatus.Text = "Nenhuma alteracao."
            }
        } catch {
            $lblSaveStatus.ForeColor = $Red
            $lblSaveStatus.Text = "ERRO: $($_.Exception.Message)"
        } finally {
            $dlg.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    })

    $dlg.ShowDialog() | Out-Null
    $dlg.Dispose()
}

# ══════════════════════════════════════════
#  EVENTOS
# ══════════════════════════════════════════
$btnLogin.Add_Click({ Do-Login })
$txtPass.Add_KeyDown({ if ($_.KeyCode -eq "Return") { Do-Login } })
$txtUser.Add_KeyDown({ if ($_.KeyCode -eq "Return") { $txtPass.Focus() } })
$txtDomain.Add_KeyDown({ if ($_.KeyCode -eq "Return") { $txtUser.Focus() } })

$btnLogout.Add_Click({ Do-Logout })
$btnRefresh.Add_Click({ Load-GPOs })
$btnNewGPO.Add_Click({ Create-NewGPO })
$btnWizard.Add_Click({ New-QuickGPO })
$btnEditGPO.Add_Click({ Edit-SelectedGPO })
$btnDeleteGPO.Add_Click({ Delete-SelectedGPO })
$btnCloneGPO.Add_Click({ Clone-SelectedGPO })
$btnExportGPO.Add_Click({ Export-SelectedGPO })
$btnImportGPO.Add_Click({ Import-GPOFromFile })

$txtSearch.Add_TextChanged({ Filter-GPOs })

# Atalhos de teclado globais
$form.Add_KeyDown({
    param($s, $e)
    if (-not $dashPanel.Visible) { return }
    if ($e.KeyCode -eq "F5") { Load-GPOs; $e.Handled = $true }
    if ($e.Control -and $e.KeyCode -eq "N") { Create-NewGPO; $e.Handled = $true }
    if ($e.Control -and $e.KeyCode -eq "W") { New-QuickGPO; $e.Handled = $true }
    if ($e.Control -and $e.KeyCode -eq "D") { Clone-SelectedGPO; $e.Handled = $true }
    if ($e.Control -and $e.KeyCode -eq "E") { Export-SelectedGPO; $e.Handled = $true }
    if ($e.Control -and $e.KeyCode -eq "I") { Import-GPOFromFile; $e.Handled = $true }
    if ($e.KeyCode -eq "Delete" -and $dgv.Focused) { Delete-SelectedGPO; $e.Handled = $true }
    if ($e.Control -and $e.KeyCode -eq "F") { $txtSearch.Focus(); $txtSearch.SelectAll(); $e.Handled = $true }
})

# Duplo-clique abre editor
$dgv.Add_CellDoubleClick({
    param($s, $e)
    if ($e.RowIndex -ge 0) {
        $id   = $dgv.Rows[$e.RowIndex].Cells["Id"].Value
        $name = $dgv.Rows[$e.RowIndex].Cells["Name"].Value
        Edit-GPO -GpoId $id -GpoName $name
    }
})

# Clique direito - menu contexto
$ctxMenu = New-Object System.Windows.Forms.ContextMenuStrip
$ctxMenu.BackColor = $BgPanel
$ctxMenu.ForeColor = $FgText

$miEdit = New-Object System.Windows.Forms.ToolStripMenuItem("Editar")
$miEdit.Add_Click({
    if ($dgv.SelectedRows.Count -gt 0) {
        $id   = $dgv.SelectedRows[0].Cells["Id"].Value
        $name = $dgv.SelectedRows[0].Cells["Name"].Value
        Edit-GPO -GpoId $id -GpoName $name
    }
})
$ctxMenu.Items.Add($miEdit) | Out-Null

$miClone = New-Object System.Windows.Forms.ToolStripMenuItem("Duplicar GPO")
$miClone.Add_Click({ Clone-SelectedGPO })
$ctxMenu.Items.Add($miClone) | Out-Null

$miExport = New-Object System.Windows.Forms.ToolStripMenuItem("Exportar GPO...")
$miExport.Add_Click({ Export-SelectedGPO })
$ctxMenu.Items.Add($miExport) | Out-Null

$ctxMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator)) | Out-Null

$miDelete = New-Object System.Windows.Forms.ToolStripMenuItem("Excluir")
$miDelete.ForeColor = $Red
$miDelete.Add_Click({ Delete-SelectedGPO })
$ctxMenu.Items.Add($miDelete) | Out-Null

$dgv.ContextMenuStrip = $ctxMenu

# Selecionar row ao clicar com botao direito
$dgv.Add_CellMouseDown({
    param($s, $e)
    if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right -and $e.RowIndex -ge 0) {
        $dgv.ClearSelection()
        $dgv.Rows[$e.RowIndex].Selected = $true
    }
})

$form.Add_Resize({
    Center-LoginPanel
    $btnLogout.Location = New-Object System.Drawing.Point(($header.ClientSize.Width - 80), 12)
})

$form.Add_Shown({
    Center-LoginPanel
    if ($txtDomain.Text) { $txtPass.Focus() } else { $txtDomain.Focus() }
    # Restaurar tamanho da janela
    if ($script:Settings.WindowWidth -and $script:Settings.WindowHeight) {
        try {
            $form.Size = New-Object System.Drawing.Size([int]$script:Settings.WindowWidth, [int]$script:Settings.WindowHeight)
            $form.StartPosition = "CenterScreen"
        } catch {}
    }
    if ($script:Settings.WindowMaximized -eq $true) { $form.WindowState = "Maximized" }
})

# Salvar configuracoes ao fechar
$form.Add_FormClosing({
    try {
        $s = @{
            WindowWidth     = $form.Size.Width
            WindowHeight    = $form.Size.Height
            WindowMaximized = ($form.WindowState -eq "Maximized")
            LastDomain      = $txtDomain.Text
        }
        Save-Settings $s
    } catch {}
})

# ══════════════════════════════════════════
#  GO
# ══════════════════════════════════════════
[System.Windows.Forms.Application]::Run($form)
