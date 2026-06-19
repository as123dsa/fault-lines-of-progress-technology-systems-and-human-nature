<#
Edge-TTS 全中文音色批量生成：普通话+东北/陕西方言+粤语+台普
自动重试3次、校验文件>0KB、失败标记跳过
#>
# 你提供完整音色清单
$VoiceAll = @(
    "zh-CN-XiaoxiaoNeural",
    "zh-CN-XiaoyiNeural",
    "zh-CN-YunjianNeural",
    "zh-CN-YunxiNeural",
    "zh-CN-YunxiaNeural",
    "zh-CN-YunyangNeural",
    "zh-CN-liaoning-XiaobeiNeural",
    "zh-CN-shaanxi-XiaoniNeural",
    "zh-HK-HiuGaaiNeural",
    "zh-HK-HiuMaanNeural",
    "zh-HK-WanLungNeural",
    "zh-TW-HsiaoChenNeural",
    "zh-TW-HsiaoYuNeural"
)

# 分语种测试文本（关键：方言/粤语/台普用对应话术，大幅减少0KB报错）
$Txt_zhCN = "大家好，这里是标准普通话试听，感受AI人声效果。"
$Txt_Dialect = "今儿天气不错，出来唠唠嗑，尝尝本地特色小吃。" #东北/陕西
$Txt_Cantonese = "今日天氣好好，一齊出去行下街，飲啖茶食個包。" #粤语
$Txt_TW = "今天天氣很棒，出門逛逛、喝杯飲料吧。" #台普

# 输出文件夹
$SaveDir = Join-Path $PWD.Path "All_Chinese_Voice_Audio"
if (-not(Test-Path $SaveDir)) {New-Item $SaveDir -ItemType Directory | Out-Null}

Write-Host "==== 开始批量生成12款中文语音 ====`n" -ForegroundColor Cyan

foreach ($voice in $VoiceAll) {
    $OutFile = Join-Path $SaveDir "$voice.mp3"
    # 匹配语种选文本
    if ($voice -match "^zh-CN-liaoning|zh-CN-shaanxi"){
        $UseText = $Txt_Dialect
    }elseif ($voice -match "^zh-HK"){
        $UseText = $Txt_Cantonese
    }elseif ($voice -match "^zh-TW"){
        $UseText = $Txt_TW
    }else{
        $UseText = $Txt_zhCN
    }

    Write-Host "正在处理：$voice" -ForegroundColor Yellow
    $ok = $false
    # 最多重试3轮
    for($r=1;$r -le 3;$r++){
        # 清理旧空文件
        if(Test-Path $OutFile){Remove-Item $OutFile -Force -ErrorAction SilentlyContinue}
        # 执行合成，屏蔽控制台报错
        edge-tts --voice $voice --text "$UseText" --write-media $OutFile 2>&1 | Out-Null
        Start-Sleep 0.8
        # 判断文件>0KB才算成功
        if((Test-Path $OutFile) -and ((Get-Item $OutFile).Length -gt 1024)){
            $ok = $true
            break
        }
    }

    if($ok){
        Write-Host "✅ 生成成功：$voice`n" -ForegroundColor Green
    }else{
        Write-Host "❌ 多次失败，接口限制作废：$voice`n" -ForegroundColor Red
        # 删除残留0kb空文件
        if(Test-Path $OutFile){Remove-Item $OutFile -Force -ErrorAction SilentlyContinue}
    }
}

Write-Host "==== 全部任务结束，音频目录：$SaveDir ====" -ForegroundColor Cyan
Start-Process $SaveDir