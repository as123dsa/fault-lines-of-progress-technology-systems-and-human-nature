<#
PPT 备注转语音 + 自动播放 + 播完翻页（Office 2019 最终稳定版）
终极修复：彻底解决PowerShell字符串插值bug导致的页码丢失问题
#>

$ErrorActionPreference = 'Stop'

# ---------------------- 1. 环境检查 + 目录准备 ----------------------
$pythonPath = (Get-Command python -ErrorAction SilentlyContinue) -or (Get-Command python3 -ErrorAction SilentlyContinue)
if (-not $pythonPath) {
    Write-Host "❌ 未找到Python，请先安装Python并添加到环境变量" -ForegroundColor Red
    Read-Host -Prompt "按回车退出"
    exit 1
}

try {
    python -c "import edge_tts" | Out-Null
}
catch {
    Write-Host "🔧 正在安装 edge_tts 依赖..." -ForegroundColor Yellow
    pip install edge_tts | Out-Null
}

$outDir = Join-Path $PWD.Path "PPT语音"
New-Item -ItemType Directory -Path $outDir -Force | Out-Null

$pptxFile = Get-ChildItem -Path $PWD.Path -Filter *.pptx | Select-Object -First 1
if (-not $pptxFile) {
    Write-Host "❌ 未找到PPTX文件" -ForegroundColor Red
    Read-Host -Prompt "按回车退出"
    exit 1
}

try {
    $pptApp = New-Object -ComObject PowerPoint.Application
    $pptApp.Visible = $true
    $presentation = $pptApp.Presentations.Open($pptxFile.FullName)

    Write-Host "`n📝 开始处理幻灯片..." -ForegroundColor Cyan

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
        }
        catch {}

        if (-not $noteText) {
            Write-Host "ℹ 第 $i 页无备注，跳过" -ForegroundColor Gray
            continue
        }

        Write-Host "✅ 第 $i 页备注：$noteText" -ForegroundColor Green

        # ====================== 终极修复：完全不用字符串插值 ======================
        # 用最原始的字符串拼接，彻底绕过PowerShell的bug
        $pageNumStr = $i.ToString()
        $fileName = "第" + $pageNumStr + "页.mp3"
        $audioPath = Join-Path $outDir $fileName

        # 生成语音
        $pythonScript = @"
import asyncio
import edge_tts
TEXT = '''$noteText'''
VOICE = "zh-CN-XiaoxiaoNeural"
OUTPUT = r"$audioPath"
async def main():
    communicate = edge_tts.Communicate(TEXT, VOICE)
    await communicate.save(OUTPUT)
asyncio.run(main())
"@

        $tempPyFile = Join-Path $env:TEMP ("edge_tts_" + $pageNumStr + ".py")
        $pythonScript | Out-File -Path $tempPyFile -Encoding utf8
        python $tempPyFile | Out-Null
        Remove-Item $tempPyFile -Force

        Write-Host "✅ 第 $i 页语音生成完成：$audioPath`n" -ForegroundColor Green

        # 插入音频
        $mediaShape = $slide.Shapes.AddMediaObject2($audioPath, $false, $true, 10, 10, 20, 20)
        $mediaShape.Fill.Transparency = 1.0
        $mediaShape.Line.Transparency = 1.0
        $mediaShape.MediaFormat.Volume = 1.0

        # 自动播放动画
        $timeLine = $slide.TimeLine.MainSequence
        for ($s = $timeLine.Count; $s -ge 1; $s--) {
            if ($timeLine.Item($s).Shape -eq $mediaShape) {
                $timeLine.Item($s).Delete()
            }
        }
        $timeLine.AddEffect($mediaShape, 10, 0, 2)
        $mediaShape.AnimationSettings.PlaySettings.HideWhileNotPlaying = $true

        # 自动翻页
        $audioDurationSec = $mediaShape.MediaFormat.Length / 1000.0
        $slide.SlideShowTransition.AdvanceOnTime = $true
        $slide.SlideShowTransition.AdvanceTime = $audioDurationSec
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

Write-Host "`n🎉 全部完成！音频保存在：$outDir" -ForegroundColor Green
Read-Host "按回车退出"