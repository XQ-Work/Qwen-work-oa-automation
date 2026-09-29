# OA 工作区业务数据备份（镜像到备份盘，带日期快照）
# 用法：右键“用 PowerShell 运行”，或任务计划器定时跑；恢复时反向 robocopy 即可
$src  = 'D:\自动备份\Qwen work\自动化（OA流程）'
$dest = 'D:\自动备份\OA数据备份'   # 换备份盘/网盘盘符改这里
$stamp = Get-Date -Format 'yyyy-MM-dd'
$snap  = Join-Path $dest ("快照\" + $stamp)

foreach ($d in @('记录', '交付', '归档', '工作', '输入', '进度总览.xlsx',
                 '系统\oa-automation\config.json', '系统\模板')) {
  robocopy (Join-Path $src $d) (Join-Path $snap $d) /E /XJ /R:1 /W:1 /NFL /NDL /NJH /NJS
}
# /E 含子目录 /XJ 不跟随junction防双拷 /R:1 失败重试1次
$latest = Join-Path $dest '最新镜像'
robocopy $snap $latest /MIR /XJ /R:1 /W:1 /NFL /NDL /NJH /NJS | Out-Null
# 只保留最近 60 份日期快照，更早的自动清理
Get-ChildItem (Join-Path $dest '快照') -Directory |
  Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-60) } |
  Remove-Item -Recurse -Force
Write-Host "备份完成：$snap"
Read-Host '按回车退出'
