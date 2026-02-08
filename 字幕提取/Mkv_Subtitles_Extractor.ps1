# ==========================================
# 脚本名称: Mkv_Subtitles_Extractor.ps1
# 功能: 自动提取 MKV 内所有字幕，支持简日、简英等多种双语识别
# ==========================================

$OutputEncoding = [System.Text.Encoding]::UTF8
$mkvFiles = Get-ChildItem -Filter *.mkv

foreach ($file in $mkvFiles) {
    Write-Host "`n正在扫描视频: $($file.Name)" -ForegroundColor Cyan
    
    $streams = ffprobe -v error -select_streams s -show_entries stream=index,codec_name:stream_tags=language,title -of csv=p=0 "$($file.FullName)"
    
    if (-not $streams) {
        Write-Host "  [提示] 未发现字幕轨，跳过。" -ForegroundColor Gray
        continue
    }

    foreach ($line in $streams) {
        $parts = $line.Split(',')
        $index = $parts[0]
        $codec = $parts[1]
        $lang  = if ($parts.Count -gt 2 -and $parts[2]) { $parts[2] } else { "" }
        $title = if ($parts.Count -gt 3 -and $parts[3]) { $parts[3] } else { "" }

        # --- 核心识别逻辑：加入简日/繁日支持 ---
        $label = ""

        # 1. 简英/繁英双语
        if ($title -match "CHS/ENG|简英|Combined|Bilingual|Dual") { $label = "CHS_ENG" }
        elseif ($title -match "CHT/ENG|繁英") { $label = "CHT_ENG" }
        
        # 2. 简日/繁日双语 (新增)
        elseif ($title -match "CHS/JP|简日|日简") { $label = "CHS_JP" }
        elseif ($title -match "CHT/JP|繁日|日繁") { $label = "CHT_JP" }
        
        # 3. 简体/中文
        elseif ($title -match "CHS|Simplified|简体|简中|简字|中文") { $label = "CHS" }
        
        # 4. 繁体
        elseif ($title -match "CHT|Traditional|繁体|繁中") { $label = "CHT" }
        
        # 5. 日语 (新增)
        elseif ($title -match "JP|JAP|Japanese|日语|日文") { $label = "JP" }
        
        # 6. 纯英文
        elseif ($title -match "ENG|English|英文") { $label = "ENG" }
        
        # 7. 兜底逻辑
        else {
            $cleanTitle = $title -replace '[\\/:*?"<>|]', '_'
            $label = "$($lang)_$($cleanTitle)".Trim('_')
            if (-not $label) { $label = "Track$index" }
        }

        # --- 后缀名自动匹配 ---
        $ext = "srt"
        switch ($codec) {
            "subrip" { $ext = "srt" }
            "ass"    { $ext = "ass" }
            "ssa"    { $ext = "ass" }
            "hdmv_pgs_subtitle" { $ext = "sup" }
            "dvd_subtitle"      { $ext = "sub" }
            "mov_text"          { $ext = "mp4.srt" }
            default  { $ext = $codec }
        }

        # 构造文件名
        $outputName = "$($file.BaseName)_$label.$ext"
        
        # 执行提取
        ffmpeg -i "$($file.FullName)" -map 0:$index -c copy "$outputName" -y -loglevel error
        
        Write-Host "  [完成] 识别为 $label -> $outputName" -ForegroundColor Green
    }
}

Write-Host "`n所有任务已处理完毕！" -ForegroundColor Magenta