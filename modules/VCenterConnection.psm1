<#
.SYNOPSIS
    VMware vCenter接続管理モジュール

.DESCRIPTION
    vCenterサーバーへのセキュアな接続、切断、認証情報管理を提供します
#>

function Get-VCenterCredential {
    <#
    .SYNOPSIS
        vCenter認証情報を取得
    
    .PARAMETER Config
        設定オブジェクト
    
    .PARAMETER Server
        vCenterサーバー名（オプション、環境変数より優先）
    
    .PARAMETER Username
        ユーザー名（オプション、環境変数より優先）
    
    .PARAMETER Password
        パスワード（オプション、環境変数より優先）
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Config,
        
        [Parameter(Mandatory = $false)]
        [string]$Server,
        
        [Parameter(Mandatory = $false)]
        [string]$Username,
        
        [Parameter(Mandatory = $false)]
        [string]$Password
    )
    
    Write-Verbose "認証情報を取得しています..."
    
    # サーバー名の決定（優先順位: パラメータ > 環境変数 > 設定ファイル）
    $vcServer = if ($Server) {
        $Server
    } elseif ($env:VCENTER_SERVER) {
        $env:VCENTER_SERVER
    } elseif ($Config.vcenter.server) {
        $Config.vcenter.server
    } else {
        throw "vCenterサーバー名が指定されていません。-Server パラメータ、VCENTER_SERVER環境変数、または設定ファイルで指定してください。"
    }
    
    # ユーザー名の決定
    $vcUsername = if ($Username) {
        $Username
    } elseif ($env:VCENTER_USERNAME) {
        $env:VCENTER_USERNAME
    } else {
        throw "ユーザー名が指定されていません。-Username パラメータまたはVCENTER_USERNAME環境変数で指定してください。"
    }
    
    # パスワードの決定
    $vcPassword = if ($Password) {
        $Password
    } elseif ($env:VCENTER_PASSWORD) {
        $env:VCENTER_PASSWORD
    } else {
        throw "パスワードが指定されていません。-Password パラメータまたはVCENTER_PASSWORD環境変数で指定してください。"
    }
    
    # PSCredentialオブジェクトの作成
    $securePassword = ConvertTo-SecureString -String $vcPassword -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($vcUsername, $securePassword)
    
    return @{
        Server = $vcServer
        Credential = $credential
        Username = $vcUsername
    }
}

function Connect-VCenterSecure {
    <#
    .SYNOPSIS
        vCenterサーバーへセキュアに接続
    
    .PARAMETER Config
        設定オブジェクト
    
    .PARAMETER CredentialInfo
        Get-VCenterCredentialから取得した認証情報
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Config,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$CredentialInfo
    )
    
    Write-Verbose "vCenterサーバーへ接続しています: $($CredentialInfo.Server)"
    
    try {
        # PowerCLIモジュールの読み込み（VMware.PowerCLI または VCF.PowerCLI）
        $powerCLIModule = $null
        
        # VCF.PowerCLIを優先的にチェック
        if (Get-Module -Name VCF.PowerCLI -ListAvailable) {
            Write-Verbose "VCF.PowerCLIモジュールを使用します"
            Import-Module VCF.PowerCLI -ErrorAction Stop
            $powerCLIModule = "VCF.PowerCLI"
        } elseif (Get-Module -Name VMware.PowerCLI -ListAvailable) {
            Write-Verbose "VMware.PowerCLIモジュールを使用します"
            Import-Module VMware.PowerCLI -ErrorAction Stop
            $powerCLIModule = "VMware.PowerCLI"
        } else {
            throw "PowerCLIモジュールがインストールされていません。`n" +
                  "以下のいずれかをインストールしてください：`n" +
                  "  Install-Module -Name VMware.PowerCLI -Scope CurrentUser`n" +
                  "  Install-Module -Name VCF.PowerCLI -Scope CurrentUser"
        }
        
        Write-Verbose "使用中のPowerCLIモジュール: $powerCLIModule"
        
        # 証明書警告の抑制
        Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false -Scope Session | Out-Null
        Set-PowerCLIConfiguration -ParticipateInCEIP $false -Confirm:$false -Scope Session | Out-Null
        
        # リトライロジックを使用した接続
        $maxAttempts = $Config.vcenter.retryAttempts
        $retryDelay = $Config.vcenter.retryDelaySeconds
        $timeout = $Config.vcenter.timeout
        
        $attempt = 1
        $connection = $null
        $lastError = $null
        
        while ($attempt -le $maxAttempts -and -not $connection) {
            try {
                Write-Verbose "接続試行 $attempt/$maxAttempts..."
                
                $connectParams = @{
                    Server = $CredentialInfo.Server
                    Credential = $CredentialInfo.Credential
                    ErrorAction = 'Stop'
                }
                
                $connection = Connect-VIServer @connectParams
                
                if ($connection) {
                    Write-Verbose "vCenterサーバーへの接続に成功しました: $($CredentialInfo.Server)"
                    Write-Verbose "vCenter バージョン: $($connection.Version) ビルド: $($connection.Build)"
                    Write-Verbose "PowerCLIモジュール: $powerCLIModule"
                    
                    return @{
                        Connection = $connection
                        Server = $CredentialInfo.Server
                        Username = $CredentialInfo.Username
                        Version = $connection.Version
                        Build = $connection.Build
                        PowerCLIModule = $powerCLIModule
                        ConnectedAt = Get-Date
                    }
                }
            } catch {
                $lastError = $_
                Write-Warning "接続試行 $attempt/$maxAttempts が失敗しました: $($_.Exception.Message)"
                
                if ($attempt -lt $maxAttempts) {
                    Write-Verbose "${retryDelay}秒後に再試行します..."
                    Start-Sleep -Seconds $retryDelay
                }
            }
            
            $attempt++
        }
        
        # すべての試行が失敗した場合
        throw "vCenterサーバーへの接続に失敗しました（$maxAttempts 回試行）: $($lastError.Exception.Message)"
        
    } catch {
        Write-Error "vCenter接続エラー: $($_.Exception.Message)"
        throw
    }
}

function Test-VCenterConnection {
    <#
    .SYNOPSIS
        vCenter接続のテスト
    
    .PARAMETER ConnectionInfo
        Connect-VCenterSecureから取得した接続情報
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$ConnectionInfo
    )
    
    try {
        $connection = $ConnectionInfo.Connection
        
        if (-not $connection) {
            return $false
        }
        
        # 簡単なクエリで接続をテスト
        $null = Get-Datacenter -Server $connection -ErrorAction Stop | Select-Object -First 1
        
        Write-Verbose "vCenter接続は有効です"
        return $true
        
    } catch {
        Write-Warning "vCenter接続テストが失敗しました: $($_.Exception.Message)"
        return $false
    }
}

function Disconnect-VCenterSafe {
    <#
    .SYNOPSIS
        vCenterサーバーから安全に切断
    
    .PARAMETER ConnectionInfo
        Connect-VCenterSecureから取得した接続情報
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$ConnectionInfo
    )
    
    try {
        if ($ConnectionInfo.Connection) {
            Write-Verbose "vCenterサーバーから切断しています: $($ConnectionInfo.Server)"
            
            Disconnect-VIServer -Server $ConnectionInfo.Connection -Confirm:$false -ErrorAction Stop
            
            $duration = ((Get-Date) - $ConnectionInfo.ConnectedAt).TotalSeconds
            Write-Verbose "vCenterサーバーから切断しました（接続時間: $([math]::Round($duration, 2))秒）"
        }
    } catch {
        Write-Warning "vCenter切断エラー: $($_.Exception.Message)"
    }
}

# モジュールメンバーのエクスポート
Export-ModuleMember -Function @(
    'Get-VCenterCredential',
    'Connect-VCenterSecure',
    'Test-VCenterConnection',
    'Disconnect-VCenterSafe'
)
