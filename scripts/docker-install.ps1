# Build and run onlyoffice-vue-demo (app only; Document Server is separate).
param(
  [string]$Image = "onlyoffice-vue-demo:latest",
  [string]$Name = "onlyoffice-vue-demo",
  [int]$Port = 19102,
  [string]$DocumentServerUrl = "http://192.168.93.128:19101",
  [string]$CallbackBaseUrl = "http://192.168.93.1:19102",
  [string]$WsBaseUrl = "ws://192.168.93.1:19102",
  [string]$JwtSecret = "+keng2vx4V2ei1k/2wAsbxjpNP/v6Ew7uhyaJ9hgOr4=",
  [string]$DocumentPath = "/files/demo.docx"
)

$ErrorActionPreference = "Stop"
Set-Location (Split-Path -Parent $PSScriptRoot)

Write-Host "Building $Image ..."
docker build -t $Image .

docker rm -f $Name 2>$null | Out-Null

Write-Host "Starting $Name on host port $Port -> container 4000 ..."
docker run -d `
  --name $Name `
  -p "${Port}:4000" `
  -e "VITE_DOCUMENT_SERVER_URL=$DocumentServerUrl" `
  -e "VITE_CALLBACK_BASE_URL=$CallbackBaseUrl" `
  -e "VITE_WS_BASE_URL=$WsBaseUrl" `
  -e "VITE_ONLYOFFICE_JWT_SECRET=$JwtSecret" `
  -e "VITE_DOCUMENT_PATH=$DocumentPath" `
  $Image

Write-Host "OK. Open $CallbackBaseUrl/"
docker logs --tail 20 $Name
