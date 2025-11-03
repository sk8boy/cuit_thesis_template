param(
    [Parameter(Mandatory = $true)]
    [string]$Version, # 例如 "v1.2.0"
    [string]$TargetBranch = "main",
    [string]$AssetPath = ".\Release\cuit_thesis_template$Version.dotm",
    [string]$ReleaseNotes = "自动发布的版本 $Version。"
)

# 1. 创建并推送 Git 标签
try {
    Write-Host "步骤 1: 创建并推送 Git 标签 $Version..." -ForegroundColor Cyan
    git tag -a $Version -m "Release $Version"
    git push origin $Version
    # git push gitee $Version
}
catch {
    Write-Host "警告: 标签创建或推送可能存在问题 - $($_.Exception.Message)" -ForegroundColor Yellow
}

# 2. 通过 API 创建 Release
$Owner = "sk8boy"
$Repo = "cuit_dissertation_template"
$AccessToken = $env:GITHUB_CUIT_PAT # 建议使用环境变量

$uri = "https://api.github.com/repos/$Owner/$Repo/releases"

$releaseParams = @{
    tag_name         = $Version
    target_commitish = $TargetBranch
    name             = "Version $Version"
    body             = $ReleaseNotes
    draft            = $false
    prerelease       = $false
}

$jsonBody = $releaseParams | ConvertTo-Json
$headers = @{ 
    "Authorization" = "token $AccessToken"
    "Accept"        = "application/vnd.github.v3+json" 
}

try {
    Write-Host "步骤 2: 通过 GitHub API 创建 Release..." -ForegroundColor Cyan
    $release = Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body $jsonBody -ContentType "application/json"
    Write-Host "✅ Release 创建成功！访问地址: $($release.html_url)" -ForegroundColor Green

    # 3. 如果有资源文件，则上传 Asset
    if (Test-Path $AssetPath) {
        Write-Host "步骤 3: 上传资源文件 $AssetPath..." -ForegroundColor Cyan
        
        # 构建上传 URL（API 返回信息中包含 upload_url 模板）
        $uploadUri = $release.upload_url -replace "\{\?name,label\}", "?name=$(Split-Path $AssetPath -Leaf)"
        
        $assetHeaders = $headers.Clone()
        $assetHeaders["Content-Type"] = "application/zip" # 根据你的文件类型调整

        # 读取文件内容并上传
        $fileBytes = [System.IO.File]::ReadAllBytes((Resolve-Path $AssetPath))
        $assetResponse = Invoke-RestMethod -Uri $uploadUri -Method Post -Headers $assetHeaders -Body $fileBytes

        Write-Host "✅ 资源文件上传成功！" -ForegroundColor Green
    }
    else {
        Write-Host "信息: 未找到资源文件 $AssetPath，跳过上传步骤。" -ForegroundColor Yellow
    }
}
catch {
    Write-Host "❌ 发布过程失败: $($_.Exception.Message)" -ForegroundColor Red
    if ($_.ErrorDetails.Message) {
        Write-Host "详细错误: $($_.ErrorDetails.Message)" -ForegroundColor Red
    }
    exit 1
}

Write-Host "`n🎉 整个发布流程已完成！" -ForegroundColor Magenta