$Workspace = Split-Path -Parent $MyInvocation.MyCommand.Path
$env:PYTHONPATH = Join-Path $Workspace ".python_env"
$env:STREAMLIT_GLOBAL_DEVELOPMENT_MODE = "false"
$Python = "C:\Users\cmoha\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
Set-Location -LiteralPath $Workspace
& $Python -m streamlit run ".\app.py" --global.developmentMode=false --server.port=8503 --server.address=127.0.0.1 --server.headless=true --browser.gatherUsageStats=false *> (Join-Path $Workspace "streamlit-run.log")
