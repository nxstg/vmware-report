<#
.SYNOPSIS
    レポート出力フォーマットモジュール

.DESCRIPTION
    収集したデータを様々な形式（JSON、HTML、CSV）で出力します
#>

function Export-ReportToJSON {
    <#
    .SYNOPSIS
        レポートデータをJSON形式で出力
    
    .PARAMETER Data
        レポートデータ
    
    .PARAMETER OutputPath
        出力ファイルパス
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Data,
        
        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )
    
    try {
        Write-Verbose "JSON形式でレポートを出力しています: $OutputPath"
        
        $jsonOutput = $Data | ConvertTo-Json -Depth 10
        $jsonOutput | Out-File -FilePath $OutputPath -Encoding UTF8
        
        Write-Verbose "JSONレポートを保存しました: $OutputPath"
        return $OutputPath
        
    } catch {
        Write-Error "JSON出力中にエラーが発生しました: $($_.Exception.Message)"
        throw
    }
}

function Export-ReportToHTML {
    <#
    .SYNOPSIS
        レポートデータをHTML形式で出力
    
    .PARAMETER Data
        レポートデータ
    
    .PARAMETER OutputPath
        出力ファイルパス
    
    .PARAMETER Config
        設定オブジェクト
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Data,
        
        [Parameter(Mandatory = $true)]
        [string]$OutputPath,
        
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Config
    )
    
    try {
        Write-Verbose "HTML形式でレポートを出力しています: $OutputPath"
        
        $html = @"
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>VMware Report - $($Data.Metadata.Server)</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { 
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: #f5f5f5;
            padding: 20px;
        }
        .container { 
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            padding: 30px;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        h1 { 
            color: #2c3e50;
            margin-bottom: 10px;
            border-bottom: 3px solid #3498db;
            padding-bottom: 10px;
        }
        h2 { 
            color: #34495e;
            margin-top: 30px;
            margin-bottom: 15px;
            padding-left: 10px;
            border-left: 4px solid #3498db;
        }
        .metadata { 
            background: #ecf0f1;
            padding: 15px;
            border-radius: 5px;
            margin-bottom: 20px;
        }
        .metadata p { 
            margin: 5px 0;
            color: #2c3e50;
        }
        table { 
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 30px;
            font-size: 14px;
        }
        th { 
            background: #3498db;
            color: white;
            padding: 12px;
            text-align: left;
            font-weight: 600;
        }
        td { 
            padding: 10px 12px;
            border-bottom: 1px solid #ddd;
        }
        tr:hover { background: #f8f9fa; }
        .status-ok { color: #27ae60; font-weight: bold; }
        .status-warning { color: #f39c12; font-weight: bold; }
        .status-critical { color: #e74c3c; font-weight: bold; }
        .summary-grid {
            width: 100%;
            margin-bottom: 30px;
        }
        .summary-card {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            margin-bottom: 10px;
            display: inline-block;
            width: 23%;
            min-width: 200px;
            vertical-align: top;
            margin-right: 1%;
        }
        .summary-card h3 {
            font-size: 14px;
            opacity: 0.9;
            margin-bottom: 10px;
            margin-top: 0;
        }
        .summary-card .value {
            font-size: 32px;
            font-weight: bold;
            margin: 0;
        }
        .chart-placeholder {
            background: #ecf0f1;
            padding: 20px;
            text-align: center;
            border-radius: 5px;
            margin: 20px 0;
            color: #7f8c8d;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>🖥️ VMware Report</h1>
        
        <div class="metadata">
            <p><strong>vCenter Server:</strong> $($Data.Metadata.Server)</p>
            <p><strong>Version:</strong> $($Data.Metadata.VCenterVersion) (Build: $($Data.Metadata.VCenterBuild))</p>
            <p><strong>Report Time:</strong> $($Data.Metadata.CollectionTime.ToString('yyyy/MM/dd HH:mm:ss'))</p>
            <p><strong>Collection Duration:</strong> $($Data.Metadata.CollectionDuration) seconds</p>
        </div>
"@

        # サマリーカード（メールクライアント互換のテーブルレイアウト）
        if ($Data.Clusters -or $Data.Hosts -or $Data.VMs) {
            $html += @"
        <h2>📊 Summary</h2>
        <table cellpadding="0" cellspacing="10" border="0" style="width: 100%; margin-bottom: 30px;">
            <tr>
"@
            if ($Data.Clusters) {
                $html += @"
                <td style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); background-color: #667eea; color: white; padding: 20px; border-radius: 8px; width: 23%; min-width: 200px; vertical-align: top;">
                    <div style="font-size: 14px; opacity: 0.9; margin-bottom: 10px;">Clusters</div>
                    <div style="font-size: 32px; font-weight: bold;">$($Data.Clusters.Count)</div>
                </td>
"@
            }
            
            if ($Data.Hosts) {
                $html += @"
                <td style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); background-color: #667eea; color: white; padding: 20px; border-radius: 8px; width: 23%; min-width: 200px; vertical-align: top;">
                    <div style="font-size: 14px; opacity: 0.9; margin-bottom: 10px;">ESXi Hosts</div>
                    <div style="font-size: 32px; font-weight: bold;">$($Data.Hosts.Count)</div>
                </td>
"@
            }
            
            if ($Data.VMs) {
                $html += @"
                <td style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); background-color: #667eea; color: white; padding: 20px; border-radius: 8px; width: 23%; min-width: 200px; vertical-align: top;">
                    <div style="font-size: 14px; opacity: 0.9; margin-bottom: 10px;">Virtual Machines</div>
                    <div style="font-size: 32px; font-weight: bold;">$($Data.VMs.Count)</div>
                </td>
"@
            }
            
            if ($Data.Datastores) {
                $html += @"
                <td style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); background-color: #667eea; color: white; padding: 20px; border-radius: 8px; width: 23%; min-width: 200px; vertical-align: top;">
                    <div style="font-size: 14px; opacity: 0.9; margin-bottom: 10px;">Datastores</div>
                    <div style="font-size: 32px; font-weight: bold;">$($Data.Datastores.Count)</div>
                </td>
"@
            }
            
            $html += @"
            </tr>
        </table>
"@
        }

        # クラスターセクション
        if ($Data.Clusters) {
            $html += @"
        <h2>🔧 Clusters</h2>
        <table>
            <thead>
                <tr>
                    <th>Name</th>
                    <th>HA Enabled</th>
                    <th>DRS Enabled</th>
                    <th>DRS Level</th>
                    <th>Hosts</th>
                    <th>VMs</th>
                    <th>Total CPU</th>
                    <th>Total Memory (GB)</th>
                    <th>Status</th>
                </tr>
            </thead>
            <tbody>
"@
            foreach ($cluster in $Data.Clusters) {
                $statusClass = switch ($cluster.OverallStatus) {
                    "Critical" { "status-critical" }
                    "Warning" { "status-warning" }
                    default { "status-ok" }
                }
                
                $html += @"
                <tr>
                    <td><strong>$($cluster.Name)</strong></td>
                    <td>$($cluster.HAEnabled)</td>
                    <td>$($cluster.DrsEnabled)</td>
                    <td>$($cluster.DrsAutomationLevel)</td>
                    <td>$($cluster.NumHosts)</td>
                    <td>$($cluster.NumVMs)</td>
                    <td>$($cluster.TotalCpu)</td>
                    <td>$($cluster.TotalMemoryGB)</td>
                    <td class="$statusClass">$($cluster.OverallStatus)</td>
                </tr>
"@
            }
            $html += @"
            </tbody>
        </table>
"@
        }

        # ホストセクション
        if ($Data.Hosts) {
            $html += @"
        <h2>🖥️ ESXi Hosts</h2>
        <table>
            <thead>
                <tr>
                    <th>Name</th>
                    <th>Cluster</th>
                    <th>CPU Usage %</th>
                    <th>CPU Status</th>
                    <th>Memory Usage %</th>
                    <th>Memory Status</th>
                    <th>VMs</th>
                    <th>Uptime (days)</th>
                </tr>
            </thead>
            <tbody>
"@
            foreach ($esxiHost in $Data.Hosts) {
                $cpuStatusClass = switch ($esxiHost.CpuStatus) {
                    "Critical" { "status-critical" }
                    "Warning" { "status-warning" }
                    default { "status-ok" }
                }
                
                $memStatusClass = switch ($esxiHost.MemoryStatus) {
                    "Critical" { "status-critical" }
                    "Warning" { "status-warning" }
                    default { "status-ok" }
                }
                
                $html += @"
                <tr>
                    <td><strong>$($esxiHost.Name)</strong></td>
                    <td>$($esxiHost.Cluster)</td>
                    <td>$($esxiHost.CpuUsagePercent)%</td>
                    <td class="$cpuStatusClass">$($esxiHost.CpuStatus)</td>
                    <td>$($esxiHost.MemoryUsagePercent)%</td>
                    <td class="$memStatusClass">$($esxiHost.MemoryStatus)</td>
                    <td>$($esxiHost.NumVMs)</td>
                    <td>$($esxiHost.UptimeDays)</td>
                </tr>
"@
            }
            $html += @"
            </tbody>
        </table>
"@
        }

        # データストアセクション
        if ($Data.Datastores) {
            $html += @"
        <h2>💾 Datastores</h2>
        <table>
            <thead>
                <tr>
                    <th>Name</th>
                    <th>Type</th>
                    <th>Capacity (GB)</th>
                    <th>Used (GB)</th>
                    <th>Free (GB)</th>
                    <th>Used %</th>
                    <th>Provisioned (GB)</th>
                    <th>VMs</th>
                    <th>Status</th>
                </tr>
            </thead>
            <tbody>
"@
            foreach ($ds in $Data.Datastores) {
                $statusClass = switch ($ds.Status) {
                    "Critical" { "status-critical" }
                    "Warning" { "status-warning" }
                    default { "status-ok" }
                }
                
                $html += @"
                <tr>
                    <td><strong>$($ds.Name)</strong></td>
                    <td>$($ds.Type)</td>
                    <td>$($ds.CapacityGB)</td>
                    <td>$($ds.UsedSpaceGB)</td>
                    <td>$($ds.FreeSpaceGB)</td>
                    <td>$($ds.UsedPercent)%</td>
                    <td>$($ds.ProvisionedGB)</td>
                    <td>$($ds.NumVMs)</td>
                    <td class="$statusClass">$($ds.Status)</td>
                </tr>
"@
            }
            $html += @"
            </tbody>
        </table>
"@
        }

        # エラーイベントセクション
        if ($Data.ErrorEvents -and $Data.ErrorEvents.Count -gt 0) {
            $html += @"
        <h2>⚠️ Error Events (Last 24 Hours: $($Data.ErrorEvents.Count) Events)</h2>
        <table>
            <thead>
                <tr>
                    <th>Time</th>
                    <th>Event Type</th>
                    <th>Object</th>
                    <th>Message</th>
                    <th>User</th>
                </tr>
            </thead>
            <tbody>
"@
            foreach ($event in $Data.ErrorEvents) {
                $html += @"
                <tr>
                    <td>$($event.CreatedTime.ToString('yyyy/MM/dd HH:mm:ss'))</td>
                    <td>$($event.EventType)</td>
                    <td><strong>$($event.ObjectName)</strong> ($($event.ObjectType))</td>
                    <td>$($event.Message)</td>
                    <td>$($event.UserName)</td>
                </tr>
"@
            }
            $html += @"
            </tbody>
        </table>
"@
        }

        # エラータスクセクション
        if ($Data.ErrorTasks -and $Data.ErrorTasks.Count -gt 0) {
            $html += @"
        <h2>❌ Error Tasks (Last 24 Hours: $($Data.ErrorTasks.Count) Tasks)</h2>
        <table>
            <thead>
                <tr>
                    <th>Task Name</th>
                    <th>Description</th>
                    <th>Start Time</th>
                    <th>Finish Time</th>
                    <th>Object</th>
                    <th>User</th>
                </tr>
            </thead>
            <tbody>
"@
            foreach ($task in $Data.ErrorTasks) {
                $html += @"
                <tr>
                    <td><strong>$($task.Name)</strong></td>
                    <td>$($task.Description)</td>
                    <td>$($task.StartTime.ToString('yyyy/MM/dd HH:mm:ss'))</td>
                    <td>$($task.FinishTime.ToString('yyyy/MM/dd HH:mm:ss'))</td>
                    <td>$($task.ObjectName)</td>
                    <td>$($task.User)</td>
                </tr>
"@
            }
            $html += @"
            </tbody>
        </table>
"@
        }

        # VMセクション
        if ($Data.VMs) {
            $html += @"
        <h2>💻 Virtual Machines ($($Data.VMs.Count) VMs)</h2>
        <table>
            <thead>
                <tr>
                    <th>Name</th>
                    <th>Host</th>
                    <th>Power State</th>
                    <th>CPU</th>
                    <th>Memory (GB)</th>
                    <th>Disk (GB)</th>
                    <th>Guest OS</th>
                    <th>HA State</th>
                </tr>
            </thead>
            <tbody>
"@
            foreach ($vm in $Data.VMs) {
                $powerStateColor = if ($vm.PowerState -eq "PoweredOn") { "status-ok" } else { "status-warning" }
                $haStateColor = if ($vm.HAState -eq "Enabled") { "status-ok" } else { "status-warning" }
                
                $html += @"
                <tr>
                    <td><strong>$($vm.Name)</strong></td>
                    <td>$($vm.Host)</td>
                    <td class="$powerStateColor">$($vm.PowerState)</td>
                    <td>$($vm.NumCpu)</td>
                    <td>$($vm.MemoryGB)</td>
                    <td>$($vm.DiskTotalGB)</td>
                    <td>$($vm.GuestOS)</td>
                    <td class="$haStateColor">$($vm.HAState)</td>
                </tr>
"@
            }
            
            $html += @"
            </tbody>
        </table>
"@
        }

        $html += @"
        <footer style="margin-top: 40px; padding-top: 20px; border-top: 1px solid #ddd; text-align: center; color: #7f8c8d;">
            <p>Generated by VMware Report v1.0.0</p>
            <p>$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')</p>
        </footer>
    </div>
</body>
</html>
"@

        $html | Out-File -FilePath $OutputPath -Encoding UTF8
        
        Write-Verbose "HTMLレポートを保存しました: $OutputPath"
        return $OutputPath
        
    } catch {
        Write-Error "HTML出力中にエラーが発生しました: $($_.Exception.Message)"
        throw
    }
}

function Export-ReportToCSV {
    <#
    .SYNOPSIS
        レポートデータをCSV形式で出力
    
    .PARAMETER Data
        レポートデータ
    
    .PARAMETER OutputDirectory
        出力ディレクトリ
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Data,
        
        [Parameter(Mandatory = $true)]
        [string]$OutputDirectory
    )
    
    try {
        Write-Verbose "CSV形式でレポートを出力しています: $OutputDirectory"
        
        $exportedFiles = @()
        
        # ホストCSV
        if ($Data.Hosts) {
            $hostsPath = Join-Path $OutputDirectory "hosts.csv"
            $Data.Hosts | Export-Csv -Path $hostsPath -NoTypeInformation -Encoding UTF8
            $exportedFiles += $hostsPath
            Write-Verbose "ホストCSVを保存しました: $hostsPath"
        }
        
        # VM CSV
        if ($Data.VMs) {
            $vmsPath = Join-Path $OutputDirectory "vms.csv"
            $Data.VMs | Export-Csv -Path $vmsPath -NoTypeInformation -Encoding UTF8
            $exportedFiles += $vmsPath
            Write-Verbose "VM CSVを保存しました: $vmsPath"
        }
        
        # データストアCSV
        if ($Data.Datastores) {
            $datastoresPath = Join-Path $OutputDirectory "datastores.csv"
            $Data.Datastores | Export-Csv -Path $datastoresPath -NoTypeInformation -Encoding UTF8
            $exportedFiles += $datastoresPath
            Write-Verbose "データストアCSVを保存しました: $datastoresPath"
        }
        
        # クラスターCSV
        if ($Data.Clusters) {
            $clustersPath = Join-Path $OutputDirectory "clusters.csv"
            $Data.Clusters | Export-Csv -Path $clustersPath -NoTypeInformation -Encoding UTF8
            $exportedFiles += $clustersPath
            Write-Verbose "クラスターCSVを保存しました: $clustersPath"
        }
        
        # エラーイベントCSV
        if ($Data.ErrorEvents -and $Data.ErrorEvents.Count -gt 0) {
            $eventsPath = Join-Path $OutputDirectory "error-events.csv"
            $Data.ErrorEvents | Export-Csv -Path $eventsPath -NoTypeInformation -Encoding UTF8
            $exportedFiles += $eventsPath
            Write-Verbose "エラーイベントCSVを保存しました: $eventsPath"
        }
        
        # エラータスクCSV
        if ($Data.ErrorTasks -and $Data.ErrorTasks.Count -gt 0) {
            $tasksPath = Join-Path $OutputDirectory "error-tasks.csv"
            $Data.ErrorTasks | Export-Csv -Path $tasksPath -NoTypeInformation -Encoding UTF8
            $exportedFiles += $tasksPath
            Write-Verbose "エラータスクCSVを保存しました: $tasksPath"
        }
        
        Write-Verbose "CSVレポートを保存しました（$($exportedFiles.Count) ファイル）"
        return $exportedFiles
        
    } catch {
        Write-Error "CSV出力中にエラーが発生しました: $($_.Exception.Message)"
        throw
    }
}

function Export-Report {
    <#
    .SYNOPSIS
        指定された形式でレポートを出力
    
    .PARAMETER Data
        レポートデータ
    
    .PARAMETER Config
        設定オブジェクト
    
    .PARAMETER OutputFormats
        出力形式の配列（オプション、設定ファイルより優先）
    
    .PARAMETER OutputDirectory
        出力ディレクトリ（オプション、設定ファイルより優先）
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Data,
        
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Config,
        
        [Parameter(Mandatory = $false)]
        [string[]]$OutputFormats,
        
        [Parameter(Mandatory = $false)]
        [string]$OutputDirectory
    )
    
    try {
        # 出力形式の決定
        $formats = if ($OutputFormats) {
            $OutputFormats
        } else {
            $Config.report.outputFormats
        }
        
        # 出力ディレクトリの決定
        $outDir = if ($OutputDirectory) {
            $OutputDirectory
        } else {
            $Config.report.outputDirectory
        }
        
        # ディレクトリの作成
        if (-not (Test-Path $outDir)) {
            New-Item -ItemType Directory -Path $outDir -Force | Out-Null
            Write-Verbose "出力ディレクトリを作成しました: $outDir"
        }
        
        # タイムスタンプ付きファイル名
        $timestamp = if ($Config.report.includeTimestamp) {
            "-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
        } else {
            ""
        }
        
        $baseFileName = "$($Data.Metadata.Server -replace '\.', '-')$timestamp"
        
        $exportedFiles = @()
        
        foreach ($format in $formats) {
            switch ($format.ToLower()) {
                'json' {
                    $jsonPath = Join-Path $outDir "$baseFileName.json"
                    $null = Export-ReportToJSON -Data $Data -OutputPath $jsonPath
                    $exportedFiles += $jsonPath
                }
                'html' {
                    $htmlPath = Join-Path $outDir "$baseFileName.html"
                    $null = Export-ReportToHTML -Data $Data -OutputPath $htmlPath -Config $Config
                    $exportedFiles += $htmlPath
                }
                'csv' {
                    $csvDir = Join-Path $outDir "csv$timestamp"
                    if (-not (Test-Path $csvDir)) {
                        New-Item -ItemType Directory -Path $csvDir -Force | Out-Null
                    }
                    $csvFiles = Export-ReportToCSV -Data $Data -OutputDirectory $csvDir
                    $exportedFiles += $csvFiles
                }
            }
        }
        
        Write-Host "`n✅ レポートを出力しました:"
        foreach ($file in $exportedFiles) {
            Write-Host "   📄 $file"
        }
        
        return $exportedFiles
        
    } catch {
        Write-Error "レポート出力中にエラーが発生しました: $($_.Exception.Message)"
        throw
    }
}

# モジュールメンバーのエクスポート
Export-ModuleMember -Function @(
    'Export-ReportToJSON',
    'Export-ReportToHTML',
    'Export-ReportToCSV',
    'Export-Report'
)
