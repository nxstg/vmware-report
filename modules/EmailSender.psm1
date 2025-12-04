<#
.SYNOPSIS
    メール送信モジュール

.DESCRIPTION
    レポートをSMTP経由でメール送信します
#>

function Send-ReportEmail {
    <#
    .SYNOPSIS
        HTMLレポートをメールで送信
    
    .PARAMETER Config
        設定オブジェクト
    
    .PARAMETER ReportFiles
        送信するレポートファイルのパス配列
    
    .PARAMETER Data
        レポートデータ（サマリー生成用）
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Config,
        
        [Parameter(Mandatory = $true)]
        [string[]]$ReportFiles,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$Data
    )
    
    try {
        Write-Verbose "メール送信の準備をしています..."
        
        # メール設定の取得
        $emailConfig = $Config.email
        
        if (-not $emailConfig.enabled) {
            Write-Verbose "メール送信は無効化されています（設定ファイルで enabled = false）"
            return
        }
        
        # SMTP設定の取得（環境変数で上書き可能）
        $smtpServer = if ($env:SMTP_SERVER) { $env:SMTP_SERVER } else { $emailConfig.smtpServer }
        $smtpPort = if ($env:SMTP_PORT) { [int]$env:SMTP_PORT } else { $emailConfig.smtpPort }
        $from = if ($env:SMTP_FROM) { $env:SMTP_FROM } else { $emailConfig.from }
        $to = if ($env:SMTP_TO) { $env:SMTP_TO -split ',' } else { $emailConfig.to }
        
        # SMTP認証情報の取得
        $smtpCredential = $null
        if ($emailConfig.useAuthentication) {
            $smtpUsername = if ($env:SMTP_USERNAME) { $env:SMTP_USERNAME } else { $emailConfig.username }
            $smtpPassword = if ($env:SMTP_PASSWORD) { $env:SMTP_PASSWORD } else { $emailConfig.password }
            
            if ($smtpUsername -and $smtpPassword) {
                $securePassword = ConvertTo-SecureString -String $smtpPassword -AsPlainText -Force
                $smtpCredential = New-Object System.Management.Automation.PSCredential($smtpUsername, $securePassword)
                Write-Verbose "SMTP認証を使用します"
            } else {
                Write-Warning "SMTP認証が有効ですが、認証情報が不足しています"
            }
        }
        
        # 件名の生成
        $subject = $emailConfig.subject -replace '\{ServerName\}', $Data.Metadata.Server
        $subject = $subject -replace '\{Date\}', (Get-Date -Format 'yyyy/MM/dd')
        
        # サマリー情報の生成
        $summary = @"
vCenter Report

サーバー: $($Data.Metadata.Server)
レポート日時: $($Data.Metadata.CollectionTime.ToString('yyyy/MM/dd HH:mm:ss'))
収集時間: $($Data.Metadata.CollectionDuration)秒

=== サマリー ===
"@
        
        if ($Data.Clusters) {
            $summary += "`nクラスター: $($Data.Clusters.Count) 件"
        }
        if ($Data.Hosts) {
            $summary += "`nESXiホスト: $($Data.Hosts.Count) 件"
            $criticalHosts = @($Data.Hosts | Where-Object { $_.CpuStatus -eq 'Critical' -or $_.MemoryStatus -eq 'Critical' })
            if ($criticalHosts.Count -gt 0) {
                $summary += "`n  ⚠ Critical状態: $($criticalHosts.Count) 件"
            }
        }
        if ($Data.VMs) {
            $summary += "`n仮想マシン: $($Data.VMs.Count) 件"
            $poweredOffVMs = @($Data.VMs | Where-Object { $_.PowerState -ne 'PoweredOn' })
            if ($poweredOffVMs.Count -gt 0) {
                $summary += "`n  パワーオフ中: $($poweredOffVMs.Count) 件"
            }
        }
        if ($Data.Datastores) {
            $summary += "`nデータストア: $($Data.Datastores.Count) 件"
            $criticalDatastores = @($Data.Datastores | Where-Object { $_.Status -eq 'Critical' })
            if ($criticalDatastores.Count -gt 0) {
                $summary += "`n  ⚠ Critical状態: $($criticalDatastores.Count) 件"
            }
        }
        
        $summary += "`n`n詳細はレポートファイルをご確認ください。"
        
        # HTMLファイルを探す
        $htmlFile = $ReportFiles | Where-Object { $_ -match '\.html$' } | Select-Object -First 1
        
        # メール本文の決定
        $body = $summary
        $bodyAsHtml = $false
        $attachments = @()
        
        if ($emailConfig.attachReports) {
            # レポートを添付
            $attachments = $ReportFiles | Where-Object { Test-Path $_ }
            Write-Verbose "添付ファイル数: $($attachments.Count)"
        } elseif ($htmlFile -and $emailConfig.includeHtmlInBody) {
            # HTMLを本文に含める
            $body = Get-Content -Path $htmlFile -Raw -Encoding UTF8
            $bodyAsHtml = $true
            Write-Verbose "HTML本文を使用します"
        }
        
        # メールパラメータの構築
        $mailParams = @{
            From = $from
            To = $to
            Subject = $subject
            Body = $body
            SmtpServer = $smtpServer
            Port = $smtpPort
            Encoding = 'UTF8'
        }
        
        if ($bodyAsHtml) {
            $mailParams.BodyAsHtml = $true
        }
        
        if ($smtpCredential) {
            $mailParams.Credential = $smtpCredential
        }
        
        if ($emailConfig.useSSL) {
            $mailParams.UseSsl = $true
            Write-Verbose "SSL/TLSを使用します"
        }
        
        if ($attachments.Count -gt 0) {
            $mailParams.Attachments = $attachments
        }
        
        # CC/BCCの追加
        if ($emailConfig.cc -and $emailConfig.cc.Count -gt 0) {
            $mailParams.Cc = $emailConfig.cc
        }
        if ($emailConfig.bcc -and $emailConfig.bcc.Count -gt 0) {
            $mailParams.Bcc = $emailConfig.bcc
        }
        
        # .NET SmtpClientを使用してメール送信
        Write-Verbose "メールを送信しています..."
        Write-Verbose "  送信先: $($to -join ', ')"
        Write-Verbose "  SMTPサーバー: ${smtpServer}:${smtpPort}"
        
        # MailMessageオブジェクトの作成
        $mailMessage = New-Object System.Net.Mail.MailMessage
        $mailMessage.From = $from
        $mailMessage.Subject = $subject
        $mailMessage.Body = $body
        $mailMessage.IsBodyHtml = $bodyAsHtml
        
        # 送信先の追加
        foreach ($recipient in $to) {
            $mailMessage.To.Add($recipient)
        }
        
        # CC/BCCの追加
        if ($emailConfig.cc -and $emailConfig.cc.Count -gt 0) {
            foreach ($ccRecipient in $emailConfig.cc) {
                $mailMessage.CC.Add($ccRecipient)
            }
        }
        if ($emailConfig.bcc -and $emailConfig.bcc.Count -gt 0) {
            foreach ($bccRecipient in $emailConfig.bcc) {
                $mailMessage.Bcc.Add($bccRecipient)
            }
        }
        
        # 添付ファイルの追加
        if ($attachments.Count -gt 0) {
            foreach ($attachmentPath in $attachments) {
                $attachment = New-Object System.Net.Mail.Attachment($attachmentPath)
                $mailMessage.Attachments.Add($attachment)
            }
        }
        
        # SmtpClientの設定
        $smtpClient = New-Object System.Net.Mail.SmtpClient($smtpServer, $smtpPort)
        
        # 暗号化タイプの決定（環境変数で上書き可能）
        $encryptionType = if ($env:SMTP_ENCRYPTION) {
            $env:SMTP_ENCRYPTION
        } elseif ($emailConfig.encryptionType) {
            $emailConfig.encryptionType
        } else {
            # 後方互換性: useSSLプロパティがあれば使用
            if ($emailConfig.PSObject.Properties.Name -contains 'useSSL') {
                if ($emailConfig.useSSL) { "StartTLS" } else { "None" }
            } else {
                "StartTLS"  # デフォルト
            }
        }
        
        # 暗号化設定
        switch ($encryptionType.ToLower()) {
            "ssl" {
                # 明示的SSL（ポート465）
                $smtpClient.EnableSsl = $true
                Write-Verbose "暗号化: SSL/TLS（ポート ${smtpPort}）"
            }
            "starttls" {
                # STARTTLS（ポート587）
                $smtpClient.EnableSsl = $true
                Write-Verbose "暗号化: STARTTLS（ポート ${smtpPort}）"
            }
            "none" {
                # 暗号化なし（ポート25など）
                $smtpClient.EnableSsl = $false
                Write-Verbose "暗号化: なし（ポート ${smtpPort}）"
            }
            default {
                # デフォルトはSTARTTLS
                $smtpClient.EnableSsl = $true
                Write-Verbose "暗号化: STARTTLS（デフォルト、ポート ${smtpPort}）"
            }
        }
        
        if ($smtpCredential) {
            $smtpClient.Credentials = $smtpCredential.GetNetworkCredential()
        }
        
        try {
            # メール送信
            $smtpClient.Send($mailMessage)
            
            Write-Host "`n✅ メールを送信しました" -ForegroundColor Green
            Write-Host "   送信先: $($to -join ', ')" -ForegroundColor Gray
            if ($attachments.Count -gt 0) {
                Write-Host "   添付ファイル: $($attachments.Count) 件" -ForegroundColor Gray
            }
            
            return $true
            
        } finally {
            # リソースのクリーンアップ
            if ($mailMessage.Attachments.Count -gt 0) {
                foreach ($attachment in $mailMessage.Attachments) {
                    $attachment.Dispose()
                }
            }
            $mailMessage.Dispose()
            $smtpClient.Dispose()
        }
        
    } catch {
        Write-Error "メール送信中にエラーが発生しました: $($_.Exception.Message)"
        Write-Warning "メールの送信に失敗しましたが、レポート生成は完了しています"
        return $false
    }
}

function Test-EmailConfiguration {
    <#
    .SYNOPSIS
        メール設定のテスト
    
    .PARAMETER Config
        設定オブジェクト
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Config
    )
    
    try {
        $emailConfig = $Config.email
        
        if (-not $emailConfig.enabled) {
            Write-Host "メール送信: 無効" -ForegroundColor Yellow
            return $false
        }
        
        $issues = @()
        
        # 必須設定のチェック
        $smtpServer = if ($env:SMTP_SERVER) { $env:SMTP_SERVER } else { $emailConfig.smtpServer }
        $from = if ($env:SMTP_FROM) { $env:SMTP_FROM } else { $emailConfig.from }
        $to = if ($env:SMTP_TO) { $env:SMTP_TO } else { $emailConfig.to }
        
        if (-not $smtpServer) { $issues += "SMTPサーバーが設定されていません" }
        if (-not $from) { $issues += "送信元アドレスが設定されていません" }
        if (-not $to -or $to.Count -eq 0) { $issues += "送信先アドレスが設定されていません" }
        
        if ($emailConfig.useAuthentication) {
            $username = if ($env:SMTP_USERNAME) { $env:SMTP_USERNAME } else { $emailConfig.username }
            $password = if ($env:SMTP_PASSWORD) { $env:SMTP_PASSWORD } else { $emailConfig.password }
            
            if (-not $username) { $issues += "SMTP認証ユーザー名が設定されていません" }
            if (-not $password) { $issues += "SMTP認証パスワードが設定されていません" }
        }
        
        if ($issues.Count -gt 0) {
            Write-Host "`nメール設定の問題:" -ForegroundColor Yellow
            foreach ($issue in $issues) {
                Write-Host "  ⚠ $issue" -ForegroundColor Yellow
            }
            return $false
        }
        
        Write-Host "`n✓ メール設定は正常です" -ForegroundColor Green
        Write-Host "  SMTPサーバー: ${smtpServer}:$($emailConfig.smtpPort)" -ForegroundColor Gray
        Write-Host "  送信元: $from" -ForegroundColor Gray
        Write-Host "  送信先: $($to -join ', ')" -ForegroundColor Gray
        Write-Host "  認証: $(if ($emailConfig.useAuthentication) { '有効' } else { '無効' })" -ForegroundColor Gray
        Write-Host "  SSL/TLS: $(if ($emailConfig.useSSL) { '有効' } else { '無効' })" -ForegroundColor Gray
        
        return $true
        
    } catch {
        Write-Error "メール設定のテスト中にエラーが発生しました: $($_.Exception.Message)"
        return $false
    }
}

# モジュールメンバーのエクスポート
Export-ModuleMember -Function @(
    'Send-ReportEmail',
    'Test-EmailConfiguration'
)
