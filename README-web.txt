智能配箱工具（网页版 / Streamlit）使用说明

1) 安装 Python 3.10+。
2) 打开命令行进入本目录，安装依赖：
   pip install -r requirements-web.txt
3) 双击“启动网页配箱工具.bat”运行。
4) 浏览器访问： http://127.0.0.1:8501

包含文件：
- streamlit_app.py（网页前端）
- pack_planner.py（配箱核心逻辑）
- packing_parser.py（供应商list解析与模板填充）
- 启动网页配箱工具.bat（一键启动）
- requirements-web.txt（网页端依赖）
- 装箱方案公式模板.xlsx（模板，作为固定路径不可用时的本地回退模板）
