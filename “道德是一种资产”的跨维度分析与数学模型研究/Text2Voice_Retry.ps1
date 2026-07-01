<#
PPT备注语音补救工具
功能：手动输入文本，生成对应MP3语音，用于修复原PPT脚本生成失败的页面
保留重试机制 + 原音色 + 统一输出目录
#>
$ErrorActionPreference = 'Stop'

# ==================== 配置项（可自行修改）====================
$maxRetries = 3          # 最大重试次数
$retryDelay = 2          # 重试间隔(秒)
$minAudioSize = 1024     # 有效音频最小字节
$voiceName = "zh-CN-YunjianNeural" # 和原脚本一致的语音
# ===========================================================

# 1. 环境校验
$pythonPath = (Get-Command python -ErrorAction SilentlyContinue) -or (Get-Command python3 -ErrorAction SilentlyContinue)
if (-not $pythonPath) {
    Write-Host "❌ 未找到Python，请先安装并配置环境变量" -ForegroundColor Red
    Read-Host "按回车退出"
    exit 1
}

try {
    python -c "import edge_tts" | Out-Null
}
catch {
    Write-Host "🔧 自动安装 edge_tts 依赖..." -ForegroundColor Yellow
    pip install edge_tts | Out-Null
}

# 统一输出目录（和原脚本一致：当前目录下 PPT语音 文件夹）
$outDir = Join-Path $PWD.Path "PPT语音"
New-Item -ItemType Directory -Path $outDir -Force | Out-Null
Write-Host "📂 音频输出目录：$outDir`n" -ForegroundColor Cyan

# 2. 交互输入
$pageNum = Read-Host "请输入幻灯片页码（如 5）"
$inputText = Read-Host "请输入该页PPT备注文本"

if ([string]::IsNullOrWhiteSpace($inputText)) {
    Write-Host "❌ 文本不能为空，退出" -ForegroundColor Red
    Read-Host
    exit 1
}

# 拼接文件名（和原脚本命名规则完全一致）
$fileName = "第$pageNum页.mp3"
$audioPath = Join-Path $outDir $fileName

# 3. 带重试的TTS生成逻辑
$retryCount = 0
$ttsSuccess = $false

while ($retryCount -lt $maxRetries -and -not $ttsSuccess) {
    try {
        # 生成临时Python脚本
        $pythonScript = @"
import asyncio
import edge_tts
TEXT = '''$inputText'''
VOICE = "$voiceName"
OUTPUT = r"$audioPath"
async def main():
    communicate = edge_tts.Communicate(TEXT, VOICE)
    await communicate.save(OUTPUT)
asyncio.run(main())
"@
        $tempPy = Join-Path $env:TEMP "text2voice_$pageNum`_$retryCount.py"
        $pythonScript | Out-File -Path $tempPy -Encoding utf8

        # 执行并检测错误
        $pyOutput = & python $tempPy 2>&1
        if ($LASTEXITCODE -ne 0) {
            throw "Python执行异常，退出码:$LASTEXITCODE"
        }

        # 校验音频文件
        if (-not (Test-Path $audioPath)) { throw "音频文件未创建" }
        $fileSize = (Get-Item $audioPath).Length
        if ($fileSize -lt $minAudioSize) { throw "音频文件过小，生成异常" }

        $ttsSuccess = $true
        Write-Host "`n✅ 生成成功！文件路径：$audioPath" -ForegroundColor Green
        Write-Host "📊 文件大小：$fileSize 字节" -ForegroundColor Cyan
    }
    catch {
        $retryCount++
        # 清理损坏文件
        if (Test-Path $audioPath) {
            Remove-Item $audioPath -Force -ErrorAction SilentlyContinue
        }

        if ($retryCount -lt $maxRetries) {
            Write-Host "`n⚠️ 第 $retryCount 次重试失败：$_" -ForegroundColor Yellow
            Write-Host "⏳ 等待 $retryDelay 秒后继续重试..."
            Start-Sleep $retryDelay
        }
        else {
            Write-Host "`n❌ 已达最大重试次数($maxRetries)，生成彻底失败" -ForegroundColor Red
        }
    }
    finally {
        # 清理临时文件
        if (Test-Path $tempPy) {
            Remove-Item $tempPy -Force -ErrorAction SilentlyContinue
        }
    }
}

Write-Host "`n----------------------------------------"
Read-Host "按回车键关闭窗口"