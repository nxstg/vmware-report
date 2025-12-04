<#
.SYNOPSIS
    VMware vCenterデータ収集モジュール

.DESCRIPTION
    vCenterからホスト、VM、データストア、クラスター情報を収集します
#>

function Get-HostMetrics {
    <#
    .SYNOPSIS
        ESXiホストのメトリクス情報を収集
    
    .PARAMETER Connection
        vCenter接続オブジェクト
    
    .PARAMETER Config
        設定オブジェクト
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Connection,
        
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Config
    )
    
    Write-Verbose "ホストメトリクスを収集しています..."
    
    try {
        $hosts = Get-VMHost -Server $Connection | Sort-Object Name
        $hostMetrics = @()
        
        foreach ($vmhost in $hosts) {
            try {
                $cluster = Get-Cluster -VMHost $vmhost -ErrorAction SilentlyContinue
                $vmsOnHost = Get-VM -Location $vmhost -ErrorAction SilentlyContinue
                
                # CPU情報
                $cpuUsagePercent = if ($vmhost.CpuTotalMhz -gt 0) {
                    [math]::Round($vmhost.CpuUsageMhz / $vmhost.CpuTotalMhz * 100, 2)
                } else { 0 }
                
                $provisionedCpu = if ($vmsOnHost) {
                    ($vmsOnHost | Measure-Object -Property NumCpu -Sum).Sum
                } else { 0 }
                
                # メモリ情報
                $totalMemoryGB = [math]::Round($vmhost.MemoryTotalGB, 2)
                $usedMemoryGB = [math]::Round($vmhost.MemoryUsageGB, 2)
                $usedMemoryPercent = if ($vmhost.MemoryTotalGB -gt 0) {
                    [math]::Round($vmhost.MemoryUsageGB / $vmhost.MemoryTotalGB * 100, 2)
                } else { 0 }
                
                $provisionedMemoryGB = if ($vmsOnHost) {
                    [math]::Round(($vmsOnHost | Measure-Object -Property MemoryGB -Sum).Sum, 2)
                } else { 0 }
                
                # ステータス判定
                $cpuStatus = if ($cpuUsagePercent -ge $Config.report.thresholds.cpu.critical) {
                    "Critical"
                } elseif ($cpuUsagePercent -ge $Config.report.thresholds.cpu.warning) {
                    "Warning"
                } else {
                    "OK"
                }
                
                $memoryStatus = if ($usedMemoryPercent -ge $Config.report.thresholds.memory.critical) {
                    "Critical"
                } elseif ($usedMemoryPercent -ge $Config.report.thresholds.memory.warning) {
                    "Warning"
                } else {
                    "OK"
                }
                
                $uptime = [math]::Round(((Get-Date) - $vmhost.ExtensionData.Runtime.BootTime).TotalDays, 1)
                
                $hostMetrics += [PSCustomObject]@{
                    Name = $vmhost.Name
                    Cluster = $cluster.Name
                    State = $vmhost.ConnectionState.ToString()
                    PowerState = $vmhost.PowerState.ToString()
                    Version = $vmhost.Version
                    Build = $vmhost.Build
                    Manufacturer = $vmhost.Manufacturer
                    Model = $vmhost.Model
                    CpuModel = $vmhost.ProcessorType
                    NumCpuCores = $vmhost.NumCpu
                    CpuTotalMhz = $vmhost.CpuTotalMhz
                    CpuUsageMhz = $vmhost.CpuUsageMhz
                    CpuUsagePercent = $cpuUsagePercent
                    CpuStatus = $cpuStatus
                    ProvisionedCpu = $provisionedCpu
                    MemoryTotalGB = $totalMemoryGB
                    MemoryUsageGB = $usedMemoryGB
                    MemoryUsagePercent = $usedMemoryPercent
                    MemoryStatus = $memoryStatus
                    ProvisionedMemoryGB = $provisionedMemoryGB
                    NumVMs = $vmsOnHost.Count
                    BootTime = $vmhost.ExtensionData.Runtime.BootTime
                    UptimeDays = $uptime
                }
                
            } catch {
                Write-Warning "ホスト $($vmhost.Name) のメトリクス収集に失敗しました: $($_.Exception.Message)"
            }
        }
        
        Write-Verbose "ホストメトリクスの収集が完了しました（$($hostMetrics.Count) 件）"
        return $hostMetrics
        
    } catch {
        Write-Error "ホストメトリクスの収集中にエラーが発生しました: $($_.Exception.Message)"
        throw
    }
}

function Get-VMMetrics {
    <#
    .SYNOPSIS
        仮想マシンのメトリクス情報を収集
    
    .PARAMETER Connection
        vCenter接続オブジェクト
    
    .PARAMETER Config
        設定オブジェクト
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Connection,
        
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Config
    )
    
    Write-Verbose "VMメトリクスを収集しています..."
    
    try {
        $vms = Get-VM -Server $Connection | Sort-Object Name
        $vmMetrics = @()
        
        foreach ($vm in $vms) {
            try {
                $vmHost = Get-VMHost -VM $vm -ErrorAction SilentlyContinue
                $resourcePool = Get-ResourcePool -VM $vm -ErrorAction SilentlyContinue
                $disks = Get-HardDisk -VM $vm -ErrorAction SilentlyContinue
                
                $totalDiskGB = if ($disks) {
                    [math]::Round(($disks | Measure-Object -Property CapacityGB -Sum).Sum, 2)
                } else { 0 }
                
                $haState = if ($vm.HARestartPriority -eq "Disabled") {
                    "Disabled"
                } else {
                    "Enabled"
                }
                
                $vmType = "Type-X$($vm.NumCpu)M$([math]::Round($vm.MemoryGB))"
                
                $vmMetrics += [PSCustomObject]@{
                    Name = $vm.Name
                    Host = $vmHost.Name
                    ResourcePool = $resourcePool.Name
                    PowerState = $vm.PowerState.ToString()
                    NumCpu = $vm.NumCpu
                    MemoryGB = [math]::Round($vm.MemoryGB, 2)
                    MemoryMB = $vm.MemoryMB
                    DiskTotalGB = $totalDiskGB
                    DiskCount = $disks.Count
                    GuestOS = $vm.ExtensionData.Config.GuestFullName
                    GuestOSShort = $vm.Guest.OSFullName
                    VMwareToolsVersion = $vm.ExtensionData.Config.Tools.ToolsVersion
                    VMwareToolsStatus = $vm.ExtensionData.Guest.ToolsStatus
                    HARestartPriority = $vm.HARestartPriority
                    HAState = $haState
                    VMType = $vmType
                    Folder = $vm.Folder.Name
                    Notes = $vm.Notes
                }
                
            } catch {
                Write-Warning "VM $($vm.Name) のメトリクス収集に失敗しました: $($_.Exception.Message)"
            }
        }
        
        Write-Verbose "VMメトリクスの収集が完了しました（$($vmMetrics.Count) 件）"
        return $vmMetrics
        
    } catch {
        Write-Error "VMメトリクスの収集中にエラーが発生しました: $($_.Exception.Message)"
        throw
    }
}

function Get-DatastoreMetrics {
    <#
    .SYNOPSIS
        データストアのメトリクス情報を収集
    
    .PARAMETER Connection
        vCenter接続オブジェクト
    
    .PARAMETER Config
        設定オブジェクト
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Connection,
        
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Config
    )
    
    Write-Verbose "データストアメトリクスを収集しています..."
    
    try {
        $datastores = Get-Datastore -Server $Connection | Sort-Object Name
        $datastoreMetrics = @()
        
        foreach ($datastore in $datastores) {
            try {
                $vmsOnDatastore = Get-VM -Datastore $datastore -ErrorAction SilentlyContinue
                
                $capacityGB = [math]::Round($datastore.CapacityGB, 2)
                $freeSpaceGB = [math]::Round($datastore.FreeSpaceGB, 2)
                $usedSpaceGB = [math]::Round($capacityGB - $freeSpaceGB, 2)
                $usedPercent = if ($capacityGB -gt 0) {
                    [math]::Round(($usedSpaceGB / $capacityGB) * 100, 2)
                } else { 0 }
                
                # プロビジョニング容量の計算
                $provisionedGB = 0
                if ($vmsOnDatastore) {
                    $disks = $vmsOnDatastore | Get-HardDisk -ErrorAction SilentlyContinue
                    if ($disks) {
                        $provisionedGB = [math]::Round(($disks | Measure-Object -Property CapacityGB -Sum).Sum, 2)
                    }
                }
                
                # ステータス判定
                $status = if ($usedPercent -ge $Config.report.thresholds.datastore.critical) {
                    "Critical"
                } elseif ($usedPercent -ge $Config.report.thresholds.datastore.warning) {
                    "Warning"
                } else {
                    "OK"
                }
                
                $datastoreMetrics += [PSCustomObject]@{
                    Name = $datastore.Name
                    Type = $datastore.Type
                    FileSystemVersion = $datastore.FileSystemVersion
                    CapacityGB = $capacityGB
                    FreeSpaceGB = $freeSpaceGB
                    UsedSpaceGB = $usedSpaceGB
                    UsedPercent = $usedPercent
                    ProvisionedGB = $provisionedGB
                    Status = $status
                    NumVMs = $vmsOnDatastore.Count
                    Accessible = $datastore.State -eq 'Available'
                    State = $datastore.State
                }
                
            } catch {
                Write-Warning "データストア $($datastore.Name) のメトリクス収集に失敗しました: $($_.Exception.Message)"
            }
        }
        
        Write-Verbose "データストアメトリクスの収集が完了しました（$($datastoreMetrics.Count) 件）"
        return $datastoreMetrics
        
    } catch {
        Write-Error "データストアメトリクスの収集中にエラーが発生しました: $($_.Exception.Message)"
        throw
    }
}

function Get-ClusterMetrics {
    <#
    .SYNOPSIS
        クラスターのメトリクス情報を収集
    
    .PARAMETER Connection
        vCenter接続オブジェクト
    
    .PARAMETER Config
        設定オブジェクト
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Connection,
        
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Config
    )
    
    Write-Verbose "クラスターメトリクスを収集しています..."
    
    try {
        $clusters = Get-Cluster -Server $Connection | Sort-Object Name
        $clusterMetrics = @()
        
        foreach ($cluster in $clusters) {
            try {
                $hostsInCluster = Get-VMHost -Location $cluster -ErrorAction SilentlyContinue
                $vmsInCluster = Get-VM -Location $cluster -ErrorAction SilentlyContinue
                
                # 全体的なステータス判定
                $overallStatus = "OK"
                if ($cluster.ExtensionData.OverallStatus -eq 'red') {
                    $overallStatus = "Critical"
                } elseif ($cluster.ExtensionData.OverallStatus -eq 'yellow') {
                    $overallStatus = "Warning"
                }
                
                $clusterMetrics += [PSCustomObject]@{
                    Name = $cluster.Name
                    HAEnabled = $cluster.HAEnabled
                    HAAdmissionControlEnabled = $cluster.HAAdmissionControlEnabled
                    HAFailoverLevel = $cluster.HAFailoverLevel
                    DrsEnabled = $cluster.DrsEnabled
                    DrsAutomationLevel = $cluster.DrsAutomationLevel.ToString()
                    EVCMode = $cluster.EVCMode
                    NumHosts = $hostsInCluster.Count
                    NumVMs = $vmsInCluster.Count
                    TotalCpu = ($hostsInCluster | Measure-Object -Property NumCpu -Sum).Sum
                    TotalMemoryGB = [math]::Round(($hostsInCluster | Measure-Object -Property MemoryTotalGB -Sum).Sum, 2)
                    OverallStatus = $overallStatus
                    ConfigStatus = $cluster.ExtensionData.ConfigStatus
                }
                
            } catch {
                Write-Warning "クラスター $($cluster.Name) のメトリクス収集に失敗しました: $($_.Exception.Message)"
            }
        }
        
        Write-Verbose "クラスターメトリクスの収集が完了しました（$($clusterMetrics.Count) 件）"
        return $clusterMetrics
        
    } catch {
        Write-Error "クラスターメトリクスの収集中にエラーが発生しました: $($_.Exception.Message)"
        throw
    }
}

function Get-ErrorEvents {
    <#
    .SYNOPSIS
        エラーイベントを収集
    
    .PARAMETER Connection
        vCenter接続オブジェクト
    
    .PARAMETER Config
        設定オブジェクト
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Connection,
        
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Config
    )
    
    Write-Verbose "エラーイベントを収集しています..."
    
    try {
        # 過去24時間のエラーイベントを取得
        $startTime = (Get-Date).AddHours(-24)
        
        $events = Get-VIEvent -Server $Connection -Start $startTime -MaxSamples 1000 -Types Error | 
            Sort-Object CreatedTime -Descending
        
        $errorEvents = @()
        
        foreach ($event in $events) {
            try {
                $errorEvents += [PSCustomObject]@{
                    CreatedTime = $event.CreatedTime
                    EventType = $event.GetType().Name
                    Message = $event.FullFormattedMessage
                    ObjectName = $event.ObjectName
                    ObjectType = $event.ObjectType
                    UserName = $event.UserName
                    ComputerName = $event.ComputerName
                    Datacenter = $event.Datacenter.Name
                    ChainId = $event.ChainId
                }
            } catch {
                Write-Warning "イベント処理に失敗しました: $($_.Exception.Message)"
            }
        }
        
        Write-Verbose "エラーイベントの収集が完了しました（$($errorEvents.Count) 件）"
        return $errorEvents
        
    } catch {
        Write-Error "エラーイベントの収集中にエラーが発生しました: $($_.Exception.Message)"
        return @()
    }
}

function Get-ErrorTasks {
    <#
    .SYNOPSIS
        エラータスクを収集
    
    .PARAMETER Connection
        vCenter接続オブジェクト
    
    .PARAMETER Config
        設定オブジェクト
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Connection,
        
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Config
    )
    
    Write-Verbose "エラータスクを収集しています..."
    
    try {
        # 過去24時間のタスクを取得（時間フィルタリングは後で実施）
        $startTime = (Get-Date).AddHours(-24)
        
        # エラー状態のタスクを取得
        $tasks = Get-Task -Server $Connection -Status Error -ErrorAction SilentlyContinue
        
        # 過去24時間のタスクにフィルタリング
        if ($tasks) {
            $tasks = $tasks | Where-Object { $_.StartTime -ge $startTime } | Sort-Object StartTime -Descending
        }
        
        $errorTasks = @()
        
        foreach ($task in $tasks) {
            try {
                $errorTasks += [PSCustomObject]@{
                    Name = $task.Name
                    Description = $task.Description
                    State = $task.State
                    StartTime = $task.StartTime
                    FinishTime = $task.FinishTime
                    Result = $task.Result
                    ObjectName = $task.ObjectId
                    PercentComplete = $task.PercentComplete
                    IsCancellable = $task.IsCancellable
                    User = $task.User
                }
            } catch {
                Write-Warning "タスク処理に失敗しました: $($_.Exception.Message)"
            }
        }
        
        Write-Verbose "エラータスクの収集が完了しました（$($errorTasks.Count) 件）"
        return $errorTasks
        
    } catch {
        Write-Warning "エラータスクの収集中にエラーが発生しました: $($_.Exception.Message)"
        return @()
    }
}

function Invoke-DataCollection {
    <#
    .SYNOPSIS
        すべてのメトリクスを収集
    
    .PARAMETER ConnectionInfo
        vCenter接続情報
    
    .PARAMETER Config
        設定オブジェクト
    
    .PARAMETER Sections
        収集するセクション（オプション、設定ファイルより優先）
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$ConnectionInfo,
        
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Config,
        
        [Parameter(Mandatory = $false)]
        [string[]]$Sections
    )
    
    $startTime = Get-Date
    Write-Verbose "データ収集を開始します..."
    
    # 収集するセクションの決定
    $sectionsToCollect = if ($Sections) {
        $Sections
    } else {
        $Config.report.sections
    }
    
    $collectedData = @{
        Metadata = @{
            CollectionTime = $startTime
            Server = $ConnectionInfo.Server
            VCenterVersion = $ConnectionInfo.Version
            VCenterBuild = $ConnectionInfo.Build
            Sections = $sectionsToCollect
        }
    }
    
    try {
        foreach ($section in $sectionsToCollect) {
            switch ($section.ToLower()) {
                'host' {
                    $collectedData.Hosts = Get-HostMetrics -Connection $ConnectionInfo.Connection -Config $Config
                }
                'cpu' {
                    if (-not $collectedData.Hosts) {
                        $collectedData.Hosts = Get-HostMetrics -Connection $ConnectionInfo.Connection -Config $Config
                    }
                }
                'memory' {
                    if (-not $collectedData.Hosts) {
                        $collectedData.Hosts = Get-HostMetrics -Connection $ConnectionInfo.Connection -Config $Config
                    }
                }
                'vm' {
                    $collectedData.VMs = Get-VMMetrics -Connection $ConnectionInfo.Connection -Config $Config
                }
                'datastore' {
                    $collectedData.Datastores = Get-DatastoreMetrics -Connection $ConnectionInfo.Connection -Config $Config
                }
                'cluster' {
                    $collectedData.Clusters = Get-ClusterMetrics -Connection $ConnectionInfo.Connection -Config $Config
                }
                'events' {
                    $collectedData.ErrorEvents = Get-ErrorEvents -Connection $ConnectionInfo.Connection -Config $Config
                }
                'tasks' {
                    $collectedData.ErrorTasks = Get-ErrorTasks -Connection $ConnectionInfo.Connection -Config $Config
                }
            }
        }
        
        # エラーイベントとタスクは常に収集（セクションに関係なく）
        if (-not $collectedData.ErrorEvents) {
            $collectedData.ErrorEvents = Get-ErrorEvents -Connection $ConnectionInfo.Connection -Config $Config
        }
        if (-not $collectedData.ErrorTasks) {
            $collectedData.ErrorTasks = Get-ErrorTasks -Connection $ConnectionInfo.Connection -Config $Config
        }
        
        $duration = ((Get-Date) - $startTime).TotalSeconds
        $collectedData.Metadata.CollectionDuration = [math]::Round($duration, 2)
        
        Write-Verbose "データ収集が完了しました（所要時間: $duration 秒）"
        return $collectedData
        
    } catch {
        Write-Error "データ収集中にエラーが発生しました: $($_.Exception.Message)"
        throw
    }
}

# モジュールメンバーのエクスポート
Export-ModuleMember -Function @(
    'Get-HostMetrics',
    'Get-VMMetrics',
    'Get-DatastoreMetrics',
    'Get-ClusterMetrics',
    'Get-ErrorEvents',
    'Get-ErrorTasks',
    'Invoke-DataCollection'
)
