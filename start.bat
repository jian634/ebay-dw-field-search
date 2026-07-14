@echo off
cd /d "%~dp0"

:: 和汽配工具一样，二选一设置 key
:: HubGPT (推荐，支持 Claude):
::   set HUBGPT_API_KEY=你的IAF_token
::   set LLM_MODEL=claude-haiku-4-5
::
:: SiliconFlow (DeepSeek-V3):
::   set SF_API_KEY=sk-...

echo Starting eBay DW Field Search Tool...
echo Open http://localhost:8082
python server.py
pause
