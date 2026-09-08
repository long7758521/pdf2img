# 一键构建脚本：安装依赖 + 安装本工具包 + 打包 pdf2img.exe
# 产物输出目录：dist\pdf2img.exe

$ErrorActionPreference = "Stop"

Write-Host "==> 安装运行依赖..." -ForegroundColor Cyan
python -m pip install -r requirements.txt

Write-Host "==> 以可编辑模式安装 pdf2img 包（使 PyInstaller 可收集）..." -ForegroundColor Cyan
python -m pip install -e .

Write-Host "==> 使用 PyInstaller 打包单文件 exe..." -ForegroundColor Cyan
python -m PyInstaller `
    --onefile `
    --console `
    --clean `
    --noconfirm `
    --name pdf2img `
    --exclude-module torch `
    --exclude-module torchvision `
    --exclude-module librosa `
    --exclude-module sklearn `
    --exclude-module lxml `
    --exclude-module cv2 `
    --exclude-module pandas `
    --exclude-module numpy `
    main.py

Write-Host "==> 打包完成！产物位于 dist\pdf2img.exe" -ForegroundColor Green