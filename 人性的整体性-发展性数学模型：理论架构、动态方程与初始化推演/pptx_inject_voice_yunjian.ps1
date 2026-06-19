<#
PPT 备注转语音 + 自动播放 + 播完翻页（Office 2019 最终稳定版）
终极修复：彻底解决PowerShell字符串插值bug导致的页码丢失问题
新增功能：TTS语音生成失败自动重试（大批量生成防失败）
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

# ---------------------- 新增：重试配置 ----------------------
$maxRetries = 3          # 最大重试次数（可根据需求调整）
$retryDelay = 2          # 每次重试间隔秒数（避免高频请求被限制）
$minAudioSize = 1024     # 最小音频文件大小（字节），防止生成空文件

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
        $pageNumStr = $i.ToString()
        $fileName = "第" + $pageNumStr + "页.mp3"
        $audioPath = Join-Path $outDir $fileName

        # ---------------------- 核心改动：带重试的TTS生成 ----------------------
        $retryCount = 0
        $ttsSuccess = $false
        while ($retryCount -lt $maxRetries -and -not $ttsSuccess) {
            try {
                # 生成Python TTS脚本
                $pythonScript = @"
import asyncio
import edge_tts
TEXT = '''$noteText'''
VOICE = "zh-CN-YunjianNeural"
OUTPUT = r"$audioPath"
async def main():
    communicate = edge_tts.Communicate(TEXT, VOICE)
    await communicate.save(OUTPUT)
asyncio.run(main())
"@

                # 临时文件加重试计数，避免并发冲突
                $tempPyFile = Join-Path $env:TEMP ("edge_tts_" + $pageNumStr + "_retry_$retryCount.py")
                $pythonScript | Out-File -Path $tempPyFile -Encoding utf8

                # 执行Python脚本并捕获错误（关键：不忽略输出，检测退出码）
                $pythonOutput = & python $tempPyFile 2>&1
                if ($LASTEXITCODE -ne 0) {
                    throw "Python执行失败（退出码：$LASTEXITCODE），错误信息：`n$pythonOutput"
                }

                # 校验：音频文件是否存在 + 大小是否合法（防止空文件）
                if (-not (Test-Path $audioPath)) {
                    throw "音频文件未生成：$audioPath"
                }
                $audioSize = (Get-Item $audioPath).Length
                if ($audioSize -lt $minAudioSize) {
                    throw "音频文件过小（仅$audioSize字节），判定生成失败"
                }

                # 所有校验通过，标记成功
                $ttsSuccess = $true
                Write-Host "✅ 第 $i 页语音生成完成（重试次数：$retryCount）：$audioPath`n" -ForegroundColor Green
            }
            catch {
                $retryCount++
                # 清理可能生成的残缺音频文件
                if (Test-Path $audioPath) {
                    Remove-Item $audioPath -Force -ErrorAction SilentlyContinue
                }
                # 输出重试提示
                if ($retryCount -lt $maxRetries) {
                    Write-Host "⚠️ 第 $i 页语音生成失败（重试$retryCount/$maxRetries），错误：$_" -ForegroundColor Yellow
                    Write-Host "⏳ 等待$retryDelay秒后重试..." -ForegroundColor Cyan
                    Start-Sleep -Seconds $retryDelay
                }
                else {
                    Write-Host "❌ 第 $i 页语音生成重试$maxRetries次仍失败，跳过该页" -ForegroundColor Red
                }
            }
            finally {
                # 清理临时Python文件（无论成败）
                if (Test-Path $tempPyFile) {
                    Remove-Item $tempPyFile -Force -ErrorAction SilentlyContinue
                }
            }
        }

        # 如果TTS最终失败，跳过后续的音频插入/翻页配置
        if (-not $ttsSuccess) {
            continue
        }

        # ---------------------- 原有逻辑：插入音频 + 自动播放 + 翻页 ----------------------
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
    Write-Host "`n❌ 全局错误：$_" -ForegroundColor Red
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
Read-Host -Prompt "按回车退出"