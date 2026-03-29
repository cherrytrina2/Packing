import subprocess
import sys
import os
from flask import Flask, redirect

app = Flask(__name__)

# 启动 Streamlit 进程
def start_streamlit():
    streamlit_script = os.path.join(os.path.dirname(os.path.dirname(__file__)), "streamlit_app.py")
    cmd = [
        sys.executable,
        "-m",
        "streamlit",
        "run",
        streamlit_script,
        "--server.port=8000",
        "--server.headless=true",
        "--server.enableCORS=false",
        "--server.enableXsrfProtection=false"
    ]
    subprocess.Popen(cmd)

# 首次访问时启动 Streamlit 并跳转
@app.route('/')
def index():
    start_streamlit()
    return redirect("http://localhost:8000")

if __name__ == "__main__":
    app.run(port=3000)
