# VMware Report Generator v1.2.0

[![Code Quality](https://github.com/nxstg/vmware-report/actions/workflows/code-quality.yml/badge.svg)](https://github.com/nxstg/vmware-report/actions/workflows/code-quality.yml)
[![Security Scan](https://github.com/nxstg/vmware-report/actions/workflows/security-scan.yml/badge.svg)](https://github.com/nxstg/vmware-report/actions/workflows/security-scan.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

vCenter環境のレポートを生成するPowerShellスクリプトです。
レポートは複数の出力形式（JSON、HTML、CSV）をサポートしています。

## 🎯 主な特徴

- ✅ **リソース監視** - CPU、メモリ、データストアの閾値ベースの状態判定
- ✅ **複数出力形式** - JSON、HTML、CSV形式でのレポート出力
- ✅ **メール送信機能** - HTMLレポートの自動メール送信

## 📋 必要要件

- PowerShell 5.1 以上
- PowerCLI モジュール（以下のいずれか）
  - **VMware.PowerCLI** - 標準のVMware PowerCLI
  - **VCF.PowerCLI** - VMware Cloud Foundation PowerCLI（推奨）
- vCenter Server へのアクセス権限

### PowerCLI のインストール

**標準のVMware PowerCLI:**
```powershell
Install-Module -Name VMware.PowerCLI -Scope CurrentUser
```

**VMware Cloud Foundation PowerCLI（推奨）:**
```powershell
Install-Module -Name VCF.PowerCLI -Scope CurrentUser
```

> **Note:** VCF.PowerCLIがインストールされている場合は自動的に優先して使用されます。

## 🚀 クイックスタート

### 1. 環境変数の設定（推奨）

```powershell
# Windows
$env:VCENTER_SERVER = "vcenter.example.com"
$env:VCENTER_USERNAME = "administrator@vsphere.local"
$env:VCENTER_PASSWORD = "YourSecurePassword"

# Linux/macOS
export VCENTER_SERVER="vcenter.example.com"
export VCENTER_USERNAME="administrator@vsphere.local"
export VCENTER_PASSWORD="YourSecurePassword"
```

### 2. スクリプトの実行

```powershell
./vmware-report.ps1
```

## 📁 ディレクトリ構造

```
vmware-report/
├── vmware-report.ps1            # メインスクリプト
├── config/
│   ├── default-config.json      # デフォルト設定
│   └── config.sample.json       # サンプル設定
├── modules/
│   ├── VCenterConnection.psm1   # 接続管理モジュール
│   ├── DataCollector.psm1       # データ収集モジュール
│   ├── OutputFormatter.psm1     # 出力フォーマットモジュール
│   └── EmailSender.psm1         # メール送信モジュール
├── reports/                     # レポート出力ディレクトリ（自動生成）
└── logs/                        # ログディレクトリ（自動生成）
```

## 🔧 使用方法

### 基本的な使用例

```powershell
# デフォルト設定で実行
./vmware-report.ps1

# 詳細ログを表示
./vmware-report.ps1 -Verbose

# 特定のサーバーを指定
./vmware-report.ps1 -Server vcenter.example.com
```

### 出力形式の指定

```powershell
# JSON形式のみ
./vmware-report.ps1 -OutputFormats json

# HTMLとCSV形式
./vmware-report.ps1 -OutputFormats html,csv

# すべての形式
./vmware-report.ps1 -OutputFormats json,html,csv
```

### レポートセクションの選択

```powershell
# CPU とメモリ情報のみ
./vmware-report.ps1 -Sections cpu,memory

# データストアとVM情報
./vmware-report.ps1 -Sections datastore,vm

# すべてのセクション（デフォルト）
./vmware-report.ps1 -Sections cpu,memory,datastore,vm,cluster,host
```

### カスタム設定ファイルの使用

```powershell
# サンプルをコピーして編集
cp config/config.sample.json config/my-config.json

# カスタム設定で実行
./vmware-report.ps1 -ConfigFile ./config/my-config.json
```

### 出力ディレクトリの指定

```powershell
./vmware-report.ps1 -OutputDirectory /path/to/reports
```

## ⚙️ 設定ファイル

`config/default-config.json` または `config/config.sample.json` をベースにカスタマイズできます。

```json
{
  "vcenter": {
    "server": "",
    "timeout": 300,
    "retryAttempts": 3,
    "retryDelaySeconds": 5
  },
  "report": {
    "sections": ["cpu", "memory", "datastore", "vm", "cluster", "host"],
    "outputFormats": ["json", "html"],
    "outputDirectory": "./reports",
    "includeTimestamp": true,
    "thresholds": {
      "cpu": { "warning": 80, "critical": 90 },
      "memory": { "warning": 85, "critical": 95 },
      "datastore": { "warning": 80, "critical": 90 }
    }
  },
  "performance": {
    "enableParallelProcessing": true,
    "maxParallelThreads": 5
  },
  "logging": {
    "enabled": true,
    "level": "Info",
    "logDirectory": "./logs"
  }
}
```

### 設定項目の説明

#### vCenter 設定
- `server`: vCenterサーバー名（環境変数で上書き可能）
- `timeout`: 接続タイムアウト（秒）
- `retryAttempts`: 接続リトライ回数
- `retryDelaySeconds`: リトライ間隔（秒）

#### レポート設定
- `sections`: 収集するセクション
- `outputFormats`: 出力形式
- `outputDirectory`: 出力ディレクトリ
- `includeTimestamp`: ファイル名にタイムスタンプを含める
- `thresholds`: 各リソースの警告・クリティカル閾値（%）

#### パフォーマンス設定
- `enableParallelProcessing`: 並列処理の有効化
- `maxParallelThreads`: 最大並列スレッド数

#### ログ設定
- `enabled`: ログ出力の有効化
- `level`: ログレベル（Info, Warning, Error）
- `logDirectory`: ログディレクトリ

## 📧 メール送信機能

レポートを自動的にメールで送信できます。HTMLレポートを本文に含めるか、添付ファイルとして送信できます。

### メール設定

設定ファイル（`config/config.sample.json`）にメール設定を追加：

```json
{
  "email": {
    "enabled": true,
    "smtpServer": "smtp.example.com",
    "smtpPort": 587,
    "useSSL": true,
    "useAuthentication": true,
    "username": "report@example.com",
    "password": "",
    "from": "vmware-report@example.com",
    "to": ["admin@example.com", "team@example.com"],
    "cc": [],
    "bcc": [],
    "subject": "VMware Daily Report - {ServerName} ({Date})",
    "includeHtmlInBody": true,
    "attachReports": false
  }
}
```

### 環境変数でのメール設定（推奨）

SMTP認証情報は環境変数で管理することを推奨します：

```powershell
# Windows
$env:SMTP_SERVER = "smtp.example.com"
$env:SMTP_PORT = "587"
$env:SMTP_USERNAME = "report@example.com"
$env:SMTP_PASSWORD = "YourSMTPPassword"
$env:SMTP_FROM = "vmware-report@example.com"
$env:SMTP_TO = "admin@example.com,team@example.com"

# Linux/macOS
export SMTP_SERVER="smtp.example.com"
export SMTP_PORT="587"
export SMTP_USERNAME="report@example.com"
export SMTP_PASSWORD="YourSMTPPassword"
export SMTP_FROM="vmware-report@example.com"
export SMTP_TO="admin@example.com,team@example.com"
```

### メール設定オプション

| 設定項目 | 説明 | デフォルト |
|---------|------|-----------|
| `enabled` | メール送信の有効化 | `false` |
| `smtpServer` | SMTPサーバーアドレス | - |
| `smtpPort` | SMTPポート | `587` |
| `encryptionType` | 暗号化方式（SSL / StartTLS / None） | `StartTLS` |
| `useAuthentication` | SMTP認証を使用 | `true` |
| `username` | SMTP認証ユーザー名 | - |
| `password` | SMTP認証パスワード | - |
| `from` | 送信元アドレス | - |
| `to` | 送信先アドレス（配列） | `[]` |
| `cc` | CCアドレス（配列） | `[]` |
| `bcc` | BCCアドレス（配列） | `[]` |
| `subject` | メール件名（{ServerName}、{Date}が使用可能） | - |
| `includeHtmlInBody` | HTML本文として送信 | `true` |
| `attachReports` | レポートファイルを添付 | `false` |

### 暗号化タイプの詳細

| タイプ | 説明 | 推奨ポート | 用途 |
|--------|------|-----------|------|
| `StartTLS` | STARTTLS（推奨） | 587 | 最も一般的。平文接続後にTLS/SSLへアップグレード |
| `SSL` | 明示的SSL/TLS | 465 | 接続開始時からSSL/TLS |
| `None` | 暗号化なし | 25 | 内部ネットワーク専用（非推奨） |

### メール送信の使用例

```powershell
# メール送信を有効にして実行
./vmware-report.ps1 -ConfigFile ./config/my-config.json

# メール送信は設定ファイルで制御されます
# enabled: true にするとレポート生成後に自動送信されます
```

### メール本文のオプション

**Option 1: HTML本文として送信（推奨）**
```json
{
  "email": {
    "includeHtmlInBody": true,
    "attachReports": false
  }
}
```
- メールクライアントで直接レポートを閲覧
- 添付ファイルなし
- モバイルでも見やすい

**Option 2: 添付ファイルとして送信**
```json
{
  "email": {
    "includeHtmlInBody": false,
    "attachReports": true
  }
}
```
- テキストサマリーをメール本文に表示
- 全レポートファイル（JSON、HTML、CSV）を添付
- ファイルサイズに注意

### 一般的なSMTP設定例

**Gmail**
```json
{
  "smtpServer": "smtp.gmail.com",
  "smtpPort": 587,
  "useSSL": true,
  "useAuthentication": true
}
```
注: Gmailの場合、アプリパスワードの使用が必要です

**Office 365 / Microsoft 365**
```json
{
  "smtpServer": "smtp.office365.com",
  "smtpPort": 587,
  "useSSL": true,
  "useAuthentication": true
}
```

**AWS SES**
```json
{
  "smtpServer": "email-smtp.us-east-1.amazonaws.com",
  "smtpPort": 587,
  "useSSL": true,
  "useAuthentication": true
}
```

### トラブルシューティング（メール）

**メールが送信されない**
- `enabled: true` になっているか確認
- SMTP設定が正しいか確認
- 環境変数が設定されているか確認
- ファイアウォールでSMTPポートが開いているか確認

**認証エラー**
- SMTP認証情報が正しいか確認
- Gmailの場合はアプリパスワードを使用
- 2段階認証の設定を確認

**SSL/TLS エラー**
```json
{
  "useSSL": false
}
```
または異なるポートを試してください（25, 465, 587）

## 📊 出力形式

### JSON形式
```
reports/vcenter-example-com-20231202-143000.json
```
- プログラムでの処理に最適
- API連携やデータベース投入に使用
- すべてのデータを含む

### HTML形式
```
reports/vcenter-example-com-20231202-143000.html
```
- ブラウザで視覚的に確認
- サマリーカード、テーブル表示
- 色分けされたステータス表示

### CSV形式
```
reports/csv-20231202-143000/
├── hosts.csv
├── vms.csv
├── datastores.csv
└── clusters.csv
```
- Excelでの分析に最適
- 各リソースタイプごとに個別のCSVファイル

## 🔐 セキュリティベストプラクティス

### 1. 環境変数の使用（推奨）

```powershell
# 環境変数に認証情報を設定
$env:VCENTER_SERVER = "vcenter.example.com"
$env:VCENTER_USERNAME = "administrator@vsphere.local"
$env:VCENTER_PASSWORD = "SecurePassword"

# スクリプト実行（パスワードをコマンドラインに含めない）
./vmware-report.ps1
```

### 2. 設定ファイルでの認証情報管理

平文パスワードを設定ファイルに保存しないでください。環境変数を使用してください。

### 3. 最小権限の原則

レポート生成には読み取り専用権限で十分です。以下の権限を持つ専用アカウントを作成することを推奨します：

- **Global** → Read-only
- **Datacenter** → Read-only
- **Cluster** → Read-only
- **Host** → Read-only
- **Virtual Machine** → Read-only
- **Datastore** → Read-only

## 🔧 トラブルシューティング

### PowerCLI モジュールが見つからない

```powershell
Install-Module -Name VMware.PowerCLI -Scope CurrentUser -Force
```

### 証明書エラー

スクリプトは自動的に証明書エラーを無視しますが、手動で設定する場合：

```powershell
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false
```

### 接続タイムアウト

`config/default-config.json` の `timeout` 値を増やしてください：

```json
{
  "vcenter": {
    "timeout": 600
  }
}
```

### 権限エラー

使用しているアカウントに十分な読み取り権限があることを確認してください。

## 📝 ログ

詳細なログを表示するには `-Verbose` パラメータを使用：

```powershell
./vmware-report.ps1 -Verbose
```

ログファイルは `logs/` ディレクトリに保存されます（設定による）。

## 🧪 テスト

Pesterテストの実行：

```powershell
# Pesterのインストール
Install-Module -Name Pester -Force -SkipPublisherCheck

# テストの実行
Invoke-Pester -Path ./tests/
```

## 📚 参考資料

- [VMware PowerCLI Documentation](https://developer.vmware.com/powercli)
- [VCF PowerCLI Documentation](https://developer.vmware.com/docs/powercli-vcf)
- [vSphere API Reference](https://developer.vmware.com/apis/vsphere-automation/latest/)
- [PowerShell Documentation](https://docs.microsoft.com/powershell/)

## 🚀 リリース手順（開発者向け）

新しいバージョンをリリースする場合は、以下の手順でタグを作成してください：

```bash
# 変更をコミット
git add .
git commit -m "feat: 新機能の追加"

# バージョンタグを作成（セマンティックバージョニング）
git tag v1.0.1

# タグをプッシュ
git push origin v1.0.1
```

タグがプッシュされると、GitHub Actionsが自動的に：
1. リリースアーカイブ（ZIP）を作成
2. チェックサム（SHA256）を生成
3. 変更履歴を自動生成
4. GitHub Releasesにリリースを作成

### バージョニング規則

このプロジェクトは[セマンティックバージョニング](https://semver.org/lang/ja/)に従います：

- **MAJOR version** (v2.0.0): 互換性のない変更
- **MINOR version** (v1.1.0): 後方互換性のある機能追加
- **PATCH version** (v1.0.1): 後方互換性のあるバグ修正

## 🤝 コントリビューション

バグ報告や機能提案は Issues にお願いします。

## 📄 ライセンス

このプロジェクトは MIT ライセンスの下で公開されています。

## 👥 作者

nxstg

---

**Note:** このスクリプトは読み取り専用の操作のみを行います。vCenter環境に変更を加えることはありません。
