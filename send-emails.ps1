# craft.studio 自動メール送信スクリプト
# 使い方: .\send-emails.ps1
# ai-consulting/email-daily/ の中の「承認済み」ファイルを読んでメール送信する

# 設定読み込み
$envFile = Join-Path $PSScriptRoot ".env.local"
$gmailUser = ""
$gmailPass = ""
foreach ($line in Get-Content $envFile) {
    if ($line -match "^GMAIL_USER=(.+)") { $gmailUser = $matches[1] }
    if ($line -match "^GMAIL_APP_PASSWORD=(.+)") { $gmailPass = $matches[1] }
}

if (-not $gmailUser -or -not $gmailPass) {
    Write-Host "ERROR: .env.localが見つかりません" -ForegroundColor Red
    exit 1
}

# 今日の承認ファイルを探す
$today = Get-Date -Format "yyyy-MM-dd"
$approvedFile = Join-Path $PSScriptRoot "ai-consulting\email-daily\${today}-approved.txt"
$dailyFile = Join-Path $PSScriptRoot "ai-consulting\email-daily\${today}.md"

if (-not (Test-Path $approvedFile)) {
    Write-Host "承認ファイルが見つかりません: $approvedFile" -ForegroundColor Yellow
    Write-Host "ai-consulting/email-daily/${today}-approved.txt を作成して承認した会社番号を記入してください"
    Write-Host "例: 1,3,5,7"
    exit 0
}

if (-not (Test-Path $dailyFile)) {
    Write-Host "本日のリストが見つかりません: $dailyFile" -ForegroundColor Red
    exit 1
}

# 承認番号を読み込む
$approvedNumbers = (Get-Content $approvedFile) -split "," | ForEach-Object { $_.Trim() }
Write-Host "承認済み会社番号: $($approvedNumbers -join ', ')" -ForegroundColor Cyan

# 日次ファイルからメール情報を抽出
$content = Get-Content $dailyFile -Raw -Encoding UTF8
$companies = @()

# ### 番号. で区切られたセクションを解析
$sections = $content -split "(?=### \d+\. )"
foreach ($section in $sections) {
    if ($section -match "### (\d+)\. (.+)\n") {
        $num = $matches[1]
        $name = $matches[2].Trim()

        # メールアドレス抽出
        $email = ""
        if ($section -match "📧 連絡先: ([^\s\n]+@[^\s\n]+)") {
            $email = $matches[1]
        }

        # 件名抽出
        $subject = ""
        if ($section -match "件名: (.+)") {
            $subject = $matches[1].Trim()
        }

        # 本文抽出（```の間）
        $body = ""
        if ($section -match "```\r?\n([\s\S]+?)\r?\n```") {
            $body = $matches[1].Trim()
        }

        if ($num -and $email -and $subject -and $body) {
            $companies += @{
                Number = $num
                Name = $name
                Email = $email
                Subject = $subject
                Body = $body
            }
        }
    }
}

# 承認済みの会社にメール送信
$sent = 0
$failed = 0

foreach ($company in $companies) {
    if ($company.Number -in $approvedNumbers) {
        Write-Host "`n[$($company.Number)] $($company.Name) に送信中..." -ForegroundColor Yellow

        try {
            $smtp = New-Object System.Net.Mail.SmtpClient("smtp.gmail.com", 587)
            $smtp.EnableSsl = $true
            $smtp.Credentials = New-Object System.Net.NetworkCredential($gmailUser, $gmailPass)

            $mail = New-Object System.Net.Mail.MailMessage
            $mail.From = $gmailUser
            $mail.To.Add($company.Email)
            $mail.Subject = $company.Subject
            $mail.Body = $company.Body
            $mail.BodyEncoding = [System.Text.Encoding]::UTF8
            $mail.SubjectEncoding = [System.Text.Encoding]::UTF8

            $smtp.Send($mail)
            Write-Host "  ✅ 送信完了: $($company.Email)" -ForegroundColor Green
            $sent++
        }
        catch {
            Write-Host "  ❌ 送信失敗: $($_.Exception.Message)" -ForegroundColor Red
            $failed++
        }
    }
}

Write-Host "`n=== 送信完了 ===" -ForegroundColor Cyan
Write-Host "送信成功: $sent 件"
Write-Host "送信失敗: $failed 件"

# 承認ファイルを削除（送信済みの証拠として日付付きに変更）
if ($sent -gt 0) {
    Rename-Item $approvedFile "ai-consulting\email-daily\${today}-sent.txt"
    Write-Host "`n承認ファイルを送信済みに変更しました"
}
