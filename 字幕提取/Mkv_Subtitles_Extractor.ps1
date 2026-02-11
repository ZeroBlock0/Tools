# ==========================================
# 脚本名称: Mkv_Subtitles_Extractor.ps1
# 功能: 自动提取 MKV 内所有字幕，支持简日、简英等多种双语识别
# ==========================================

$OutputEncoding = [System.Text.Encoding]::UTF8

# ==========================================
# 依赖检查
# ==========================================
function Test-Dependency {
    param([string]$Command)
    $null = Get-Command $Command -ErrorAction SilentlyContinue
    return $?
}

if (-not (Test-Dependency "ffprobe")) {
    Write-Host "[错误] 未找到 ffprobe，请安装 FFmpeg 工具包并添加到 PATH" -ForegroundColor Red
    exit 1
}

if (-not (Test-Dependency "ffmpeg")) {
    Write-Host "[错误] 未找到 ffmpeg，请安装 FFmpeg 工具包并添加到 PATH" -ForegroundColor Red
    exit 1
}

# ==========================================
# 文件名清理函数（移除 Windows 非法字符）
# ==========================================
function Get-SafeFilename {
    param([string]$Name)
    return ($Name -replace '[\\/:*?"<>|]', '_').Trim('_').Trim()
}

# ==========================================
# 扫描 MKV 文件
# ==========================================
Write-Host "开始扫描 MKV 文件..." -ForegroundColor Cyan

# 用 @() 强制包装为数组，避免单文件时 .Count 不可用
$mkvFiles = @(Get-ChildItem -Filter *.mkv -ErrorAction SilentlyContinue)

if ($mkvFiles.Count -eq 0) {
    Write-Host "[提示] 当前目录未找到 MKV 文件" -ForegroundColor Yellow
    exit 0
}

Write-Host "找到 $($mkvFiles.Count) 个 MKV 文件" -ForegroundColor Cyan

# 统计变量
$totalSuccess = 0
$totalFail = 0
$totalSkip = 0

foreach ($file in $mkvFiles) {
    Write-Host "`n正在扫描视频: $($file.Name)" -ForegroundColor Cyan

    try {
        # 使用 JSON 格式输出以避免 CSV 逗号解析问题
        # 将 stderr 重定向到 $null 以避免与 ErrorActionPreference 冲突
        $streamsJson = & ffprobe -v error -select_streams s -show_entries "stream=index,codec_name:stream_tags=language,title" -of json "$($file.FullName)" 2>$null

        if ($LASTEXITCODE -ne 0) {
            Write-Host "  [错误] ffprobe 无法读取视频信息，跳过。" -ForegroundColor Red
            $totalSkip++
            continue
        }

        # 将字符串数组合并后再解析 JSON
        $streamData = ($streamsJson -join "`n") | ConvertFrom-Json

        if (-not $streamData.streams -or $streamData.streams.Count -eq 0) {
            Write-Host "  [提示] 未发现字幕轨，跳过。" -ForegroundColor Gray
            $totalSkip++
            continue
        }

        Write-Host "  发现 $($streamData.streams.Count) 条字幕轨" -ForegroundColor DarkCyan

        foreach ($stream in $streamData.streams) {
            $index = $stream.index
            $codec = $stream.codec_name

            # 空值安全访问 tags
            $tags  = $stream.tags
            $lang  = if ($tags -and $tags.language) { $tags.language } else { "" }
            $title = if ($tags -and $tags.title) { $tags.title } else { "" }

            # --- 核心识别逻辑：加入简日/繁日支持 ---
            $label = ""

            # 1. 简英/繁英双语
            if ($title -match "CHS[/&]ENG|简英|Combined|Bilingual|Dual") { $label = "CHS_ENG" }
            elseif ($title -match "CHT[/&]ENG|繁英") { $label = "CHT_ENG" }

            # 2. 简日/繁日双语
            elseif ($title -match "CHS[/&]JP|简日|日简") { $label = "CHS_JP" }
            elseif ($title -match "CHT[/&]JP|繁日|日繁") { $label = "CHT_JP" }

            # 3. 简体/中文（使用 \b 防止误匹配）
            elseif ($title -match "\bCHS\b|Simplified|简体|简中|简字|中文") { $label = "CHS" }

            # 4. 繁体
            elseif ($title -match "\bCHT\b|Traditional|繁体|繁中") { $label = "CHT" }

            # 5. 日语（使用 \b 防止 "JP" 误匹配其他文本）
            elseif ($title -match "\bJPN?\b|\bJAP\b|Japanese|日语|日文") { $label = "JP" }

            # 6. 纯英文
            elseif ($title -match "\bENG\b|English|英文") { $label = "ENG" }

            # 7. 兜底：使用 language tag + title
            else {
                $cleanTitle = Get-SafeFilename $title
                $label = @($lang, $cleanTitle) | Where-Object { $_ } | Join-String -Separator '_'
                if (-not $label) { $label = "Track$index" }
            }

            # 确保 label 也是安全的文件名
            $label = Get-SafeFilename $label

            # --- 后缀名自动匹配 ---
            $ext = switch ($codec) {
                "subrip"            { "srt" }
                "ass"               { "ass" }
                "ssa"               { "ass" }
                "hdmv_pgs_subtitle" { "sup" }
                "dvd_subtitle"      { "sub" }
                "mov_text"          { "srt" }
                "webvtt"            { "vtt" }
                default             { $codec }
            }

            # 构造输出完整路径，使用与源文件相同的目录
            $outputDir = $file.DirectoryName
            $outputName = "$($file.BaseName)_$label.$ext"
            $outputPath = Join-Path $outputDir $outputName

            # 自动重命名避免覆盖（使用 -LiteralPath 防止 [ ] 通配符问题）
            $counter = 1
            while (Test-Path -LiteralPath $outputPath) {
                $outputName = "$($file.BaseName)_$label($counter).$ext"
                $outputPath = Join-Path $outputDir $outputName
                $counter++
            }

            # 执行提取，stderr 单独捕获用于错误提示
            $ffmpegErr = $null
            & ffmpeg -i "$($file.FullName)" -map "0:$index" -c copy "$outputPath" -n -loglevel error 2>&1 | ForEach-Object {
                if ($_ -is [System.Management.Automation.ErrorRecord]) {
                    $ffmpegErr = $_.ToString()
                }
            }

            if ($LASTEXITCODE -eq 0) {
                Write-Host "  [完成] 轨道 $index ($codec) -> $outputName [$label]" -ForegroundColor Green
                $totalSuccess++
            } else {
                $errMsg = if ($ffmpegErr) { $ffmpegErr } else { "未知错误" }
                Write-Host "  [失败] 轨道 $index ($codec): $errMsg" -ForegroundColor Red
                $totalFail++
            }
        }
    } catch {
        Write-Host "  [错误] 处理视频时发生异常: $($_.Exception.Message)" -ForegroundColor Red
        $totalSkip++
        continue
    }
}

# ==========================================
# 统计汇总
# ==========================================
Write-Host "`n==========================================" -ForegroundColor Magenta
Write-Host "所有任务已处理完毕！" -ForegroundColor Magenta
Write-Host "  成功: $totalSuccess  失败: $totalFail  跳过: $totalSkip" -ForegroundColor Magenta
Write-Host "==========================================" -ForegroundColor Magenta
