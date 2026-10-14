# 前端启动说明 (E:\agent 开发)

# 后端 API 服务 (无新依赖, 标准库 http.server)
python -X utf8 server.py            # 默认 :8000
python -X utf8 server.py 8080      # 自定义端口

# 浏览器打开 http://127.0.0.1:8000 (或对应端口)

# 前端文件
#   index.html   单页界面: 检索 / 出题 / 批改 / 薄弱点档案 四个 tab
# 调的 API
#   GET  /            index.html
#   GET  /api/search?q=...&tk=...   M2 检索
#   GET  /api/practice?topic=&diff=&n=  M3 出题 (含薄弱点倾斜)
#   POST /api/grade   M4 批改 {mode,A|B, topic, ...}
#   GET  /api/pages/{n}   M1 页详情 (公式对象/例题/正文)
#   GET  /api/profile     薄弱点档案

# 注意: 中文 URL 参数在 PowerShell 里会被转义坏, 用浏览器或 curl 访问.
