# download-deps.ps1
# 在能联网的机器上执行，下载外部依赖

$PANDOC_VERSION = "3.1.11"
$PANDOC_URL = "https://github.com/jgm/pandoc/releases/download/${PANDOC_VERSION}/pandoc-${PANDOC_VERSION}-linux-amd64.tar.gz"
$OUT_FILE = "pandoc-${PANDOC_VERSION}-linux-amd64.tar.gz"

Write-Host "Downloading pandoc ${PANDOC_VERSION}..."
Invoke-WebRequest -Uri $PANDOC_URL -OutFile $OUT_FILE -UseBasicParsing

Write-Host "Done: ${OUT_FILE}"
Write-Host "SHA256: $((Get-FileHash $OUT_FILE -Algorithm SHA256).Hash)"