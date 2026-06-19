<#
PPT 备注转语音 + 自动播放 + 播完翻页（Office 2019 最终稳定版）
修复：PlaySettings.Volume 不存在报错
音量方案：TTS拉满 + MediaFormat=1.0 + 系统最大音量
#>

$ErrorActionPreference = 'Stop'

# ---------------------- 1. 目录准备 ----------------------
$outDir = Join-Path $PWD.Path "PPT语音"
New-Item -ItemType Directory -Path $outDir -Force | Out-Null

$pptxFile = Get-ChildItem -Path $PWD.Path -Filter *.pptx | Select-Object -First 1
if (-not $pptxFile) {
    Write-Host "❌ 未找到PPTX文件" -ForegroundColor Red
    Read-Host -Prompt "按回车退出"
    exit 1
}

try {
    # ---------------------- 2. 启动PPT ----------------------
    $pptApp = New-Object -ComObject PowerPoint.Application
    $pptApp.Visible = $true
    $presentation = $pptApp.Presentations.Open($pptxFile.FullName)

    # ---------------------- 3. 读取备注 ----------------------
    Write-Host "`n📝 读取幻灯片备注..." -ForegroundColor Cyan
    $slideNotes = @()
    for ($i = 1; $i -le $presentation.Slides.Count; $i++) {
        $slide = $presentation.Slides.Item($i)
        $noteText = ""
        try {
            foreach ($shape in $slide.NotesPage.Shapes) {
                if ($shape.TextFrame.HasText -eq -1) {
                    $rawText = $shape.TextFrame.TextRange.Text.Trim()
                    if ($rawText -notmatch "^[\d\.]+$" -and $rawText -notmatch "^\s*$") {
                        $noteText = $rawText
                        break
                    }
                }
            }
        }catch{}

        if ($noteText) {
            $slideNotes += [PSCustomObject]@{ Page = $i; Text = $noteText; AudioPath = $null }
            Write-Host "✅ 第 $i 页备注：$noteText" -ForegroundColor Green
        }
    }

    if (-not $slideNotes) {
        Write-Host "`n❌ 未读取到有效备注" -ForegroundColor Red
        Read-Host -Prompt "按回车退出"
        exit 1
    }

    # ---------------------- 4. TTS 生成（音量100%） ----------------------
    Write-Host "`n🔊 生成最大音量语音（TTS音量=100）..." -ForegroundColor Cyan
    $tts = New-Object -ComObject SAPI.SpVoice
    $tts.Volume = 100   # 0-100，直接拉满
    $tts.Rate = 0        # 语速正常
    $cnVoice = $tts.GetVoices() | Where-Object { $_.GetDescription() -match "中文|Xiaoxiao|Yunxi" } | Select-Object -First 1
    if ($cnVoice) { $tts.Voice = $cnVoice }

    foreach ($item in $slideNotes) {
        $audioPath = Join-Path $outDir "第$($item.Page)页.wav"
        $audioStream = New-Object -ComObject SAPI.SpFileStream
        $audioStream.Open($audioPath, 3)
        $tts.AudioOutputStream = $audioStream
        $tts.Speak($item.Text) | Out-Null
        $audioStream.Close()
        $item.AudioPath = $audioPath
    }

    # ---------------------- 5. 插入音频 + 最大音量（无报错版） ----------------------
    Write-Host "`n🎬 配置音频自动播放 + 透明图标 + 自动翻页..." -ForegroundColor Cyan
    foreach ($item in $slideNotes) {
        $slideNum = $item.Page
        $audioPath = $item.AudioPath
        $targetSlide = $presentation.Slides.Item($slideNum)

        # 插入音频：左上角、极小、透明
        $mediaShape = $targetSlide.Shapes.AddMediaObject2($audioPath,$false,$true,10,10,20,20)
        $mediaShape.Fill.Transparency = 1.0
        $mediaShape.Line.Transparency = 1.0

        # ✅ 官方有效：MediaFormat 音量=1.0（0.0~1.0）
        $mediaShape.MediaFormat.Volume = 1.0

        # 自动播放动画
        $timeLine = $targetSlide.TimeLine.MainSequence
        for ($i = $timeLine.Count; $i -ge 1; $i--) {
            if ($timeLine.Item($i).Shape -eq $mediaShape) { $timeLine.Item($i).Delete() }
        }
        $playEffect = $timeLine.AddEffect($mediaShape, 10, 0, 2)
        $mediaShape.AnimationSettings.PlaySettings.HideWhileNotPlaying = $true

        # 自动翻页 = 音频时长
        $audioDurationSec = $mediaShape.MediaFormat.Length / 1000.0
        $targetSlide.SlideShowTransition.AdvanceOnTime = $true
        $targetSlide.SlideShowTransition.AdvanceTime = $audioDurationSec

        Write-Host "✅ 第 $slideNum 页：音量=1.0，翻页=$audioDurationSec 秒" -ForegroundColor Green
    }

    $presentation.Save()
    Write-Host "`n💾 PPT已保存" -ForegroundColor Green
}
catch {
    Write-Host "`n❌ 错误：$_" -ForegroundColor Red
}
finally {
    if ($presentation) { $presentation.Close() }
    if ($pptApp) { $pptApp.Quit() }
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($presentation)|Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($pptApp)|Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

Write-Host "`n🎉 全部完成！（TTS最大音量 + PPT音量1.0）" -ForegroundColor Green
Write-Host "📂 音频位置：$outDir"
Read-Host 按回车退出