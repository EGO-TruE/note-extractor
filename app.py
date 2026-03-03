"""
智能笔记提取器 - 网页版
用法：python app.py
然后在浏览器（或平板）访问屏幕上显示的地址
"""

import os
import sys
import uuid
import threading
import tempfile
from pathlib import Path

from flask import Flask, request, send_file, render_template_string, redirect, url_for, jsonify, session
import main as core

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024  # 200 MB 上限

PASSCODE = os.getenv("PASSCODE", "")
app.secret_key = os.getenv("SECRET_KEY", os.urandom(24).hex())

# 内存中存储任务状态（单 worker 多线程，内存可共享）
jobs = {}

# ─────────────────────────────────────────────
# HTML 模板
# ─────────────────────────────────────────────

HTML = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>智能笔记提取器</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: -apple-system, BlinkMacSystemFont, "PingFang SC", "Segoe UI", sans-serif;
      background: #eef2f7;
      min-height: 100vh;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 20px;
    }
    .card {
      background: #fff;
      border-radius: 20px;
      padding: 40px 36px;
      max-width: 520px;
      width: 100%;
      box-shadow: 0 8px 32px rgba(0,0,0,0.10);
    }
    h1 { font-size: 22px; color: #1a1a2e; margin-bottom: 6px; }
    .subtitle { color: #888; font-size: 14px; margin-bottom: 28px; }
    .upload-area {
      border: 2px dashed #b0c8e8;
      border-radius: 14px;
      padding: 44px 20px;
      text-align: center;
      cursor: pointer;
      transition: border-color 0.2s, background 0.2s;
      margin-bottom: 14px;
      display: block;
    }
    .upload-area:hover { border-color: #4a90e2; background: #f5f9ff; }
    .upload-icon { font-size: 52px; margin-bottom: 14px; }
    .upload-text { font-size: 17px; color: #444; font-weight: 500; }
    .upload-hint { font-size: 13px; color: #aaa; margin-top: 8px; }
    input[type="file"] { display: none; }
    .selected-file {
      background: #eef6ff;
      border-radius: 8px;
      padding: 10px 14px;
      font-size: 14px;
      color: #4a90e2;
      margin-bottom: 14px;
      display: none;
      word-break: break-all;
    }
    .btn {
      display: block;
      width: 100%;
      padding: 16px;
      background: #4a90e2;
      color: #fff;
      border: none;
      border-radius: 12px;
      font-size: 16px;
      font-weight: 600;
      cursor: pointer;
      text-align: center;
      text-decoration: none;
      margin-top: 10px;
      transition: background 0.2s;
    }
    .btn:hover { background: #357abd; }
    .btn:disabled { background: #a0c0e8; cursor: not-allowed; }
    .btn-green { background: #27ae60; }
    .btn-green:hover { background: #219a52; }
    .btn-outline {
      background: #fff;
      color: #4a90e2;
      border: 2px solid #4a90e2;
    }
    .btn-outline:hover { background: #f0f7ff; }
    .status { text-align: center; padding: 16px 0 24px; }
    .spinner {
      width: 52px; height: 52px;
      border: 5px solid #e0e8f0;
      border-top-color: #4a90e2;
      border-radius: 50%;
      animation: spin 0.9s linear infinite;
      margin: 0 auto 18px;
    }
    @keyframes spin { to { transform: rotate(360deg); } }
    .status-text { font-size: 18px; color: #333; margin-bottom: 8px; font-weight: 500; }
    .status-hint { font-size: 13px; color: #aaa; }
    .done-header { font-size: 17px; color: #27ae60; font-weight: 600; margin-bottom: 18px; }
    .dl-item {
      background: #f8faff;
      border: 1px solid #d8e8f8;
      border-radius: 12px;
      padding: 18px;
      margin-bottom: 12px;
    }
    .dl-label { font-size: 12px; color: #999; margin-bottom: 4px; text-transform: uppercase; letter-spacing: 0.5px; }
    .dl-name { font-size: 14px; font-weight: 600; color: #333; margin-bottom: 12px; word-break: break-all; }
    .error-box {
      background: #fff5f5;
      border: 1px solid #ffc0c0;
      border-radius: 10px;
      padding: 16px;
      color: #c0392b;
      font-size: 14px;
      margin-bottom: 18px;
      word-break: break-all;
    }
    .login-wrap { text-align: center; padding: 8px 0 4px; }
    .login-wrap .lock { font-size: 48px; margin-bottom: 18px; }
    input[type="password"] {
      width: 100%;
      padding: 14px 16px;
      border: 1.5px solid #d0dce8;
      border-radius: 10px;
      font-size: 16px;
      margin-bottom: 14px;
      outline: none;
      transition: border-color 0.2s;
    }
    input[type="password"]:focus { border-color: #4a90e2; }
  </style>
</head>
<body>
<div class="card">
  <h1>📚 智能笔记提取器</h1>
  <p class="subtitle">上传文档，AI 自动提取重要知识点</p>

  {% if page == 'login' %}
  <div class="login-wrap">
    <div class="lock">🔒</div>
    <form action="/login" method="post">
      <input type="password" name="passcode" placeholder="请输入访问密码" autofocus>
      {% if error %}<div class="error-box">密码错误，请重试</div>{% endif %}
      <button type="submit" class="btn">进入</button>
    </form>
  </div>

  {% elif page == 'upload' %}
  <form action="/upload" method="post" enctype="multipart/form-data" id="form">
    <label class="upload-area" for="fileInput">
      <div class="upload-icon">📄</div>
      <div class="upload-text">点击选择文件</div>
      <div class="upload-hint">支持 PDF · Word · PPT · TXT · Markdown</div>
    </label>
    <input type="file" id="fileInput" name="file"
           accept=".pdf,.docx,.pptx,.txt,.md"
           onchange="onFileChange(this)">
    <div class="selected-file" id="selectedFile"></div>
    <button type="submit" class="btn" id="submitBtn">开始分析</button>
  </form>
  <script>
    function onFileChange(input) {
      var f = input.files[0];
      if (f) {
        var el = document.getElementById('selectedFile');
        el.style.display = 'block';
        el.textContent = '已选择：' + f.name;
      }
    }
    document.getElementById('form').addEventListener('submit', function() {
      var btn = document.getElementById('submitBtn');
      btn.textContent = '上传中…';
      btn.disabled = true;
    });
  </script>

  {% elif page == 'result' %}
    {% if job.status == 'processing' %}
    <div class="status">
      <div class="spinner"></div>
      <div class="status-text">AI 正在分析中…</div>
      <div class="status-hint">通常需要 30 秒至数分钟，请勿关闭此页面</div>
    </div>
    <script>
      (function poll() {
        setTimeout(function() {
          fetch('/status/{{ job_id }}')
            .then(function(r) { return r.json(); })
            .then(function(d) {
              if (d.status !== 'processing') { location.reload(); }
              else { poll(); }
            })
            .catch(function() { poll(); });
        }, 2500);
      })();
    </script>

    {% elif job.status == 'done' %}
    <p class="done-header">✅ 分析完成，共提取 {{ job.count }} 个知识点</p>
    <div>
      <div class="dl-item">
        <div class="dl-label">标注版 · 原文 + 知识点高亮</div>
        <div class="dl-name">{{ job.files.annotated_name }}</div>
        <a href="/download/{{ job_id }}/annotated" class="btn btn-green">下载标注文档</a>
      </div>
      <div class="dl-item">
        <div class="dl-label">摘要版 · 仅知识点列表</div>
        <div class="dl-name">{{ job.files.summary_name }}</div>
        <a href="/download/{{ job_id }}/summary" class="btn">下载摘要文档</a>
      </div>
    </div>
    <a href="/" class="btn btn-outline" style="margin-top:18px;">分析另一个文件</a>

    {% else %}
    <div class="error-box">❌ 处理失败：{{ job.error }}</div>
    <a href="/" class="btn btn-outline">返回重试</a>
    {% endif %}
  {% endif %}
</div>
</body>
</html>"""


# ─────────────────────────────────────────────
# 路由
# ─────────────────────────────────────────────

@app.before_request
def require_login():
    if not PASSCODE:
        return  # 未设置密码则不拦截
    if request.endpoint in ("login",):
        return
    if not session.get("auth"):
        return redirect(url_for("login"))


@app.route("/login", methods=["GET", "POST"])
def login():
    error = False
    if request.method == "POST":
        if request.form.get("passcode") == PASSCODE:
            session["auth"] = True
            return redirect(url_for("index"))
        error = True
    return render_template_string(HTML, page="login", error=error)


@app.route("/")
def index():
    return render_template_string(HTML, page="upload")


@app.route("/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return redirect(url_for("index"))

    file = request.files["file"]
    if not file.filename:
        return redirect(url_for("index"))

    original_filename = file.filename
    suffix = Path(original_filename).suffix.lower()

    # 保存上传文件到临时位置
    tmp_fd, tmp_path = tempfile.mkstemp(suffix=suffix)
    os.close(tmp_fd)
    file.save(tmp_path)

    job_id = uuid.uuid4().hex[:8]
    jobs[job_id] = {"status": "processing", "files": {}, "error": None, "count": 0}

    def process():
        try:
            text = core.read_file(Path(tmp_path))
            kps = core.extract_knowledge_points(text)

            if not kps:
                jobs[job_id].update({"status": "error", "error": "未提取到知识点，请检查文件内容是否可读"})
                return

            stem = Path(original_filename).stem
            out_dir = Path(tempfile.gettempdir()) / f"kpe_{job_id}"
            out_dir.mkdir(exist_ok=True)

            annotated_path = out_dir / f"原文_带标注_{stem}.docx"
            summary_path   = out_dir / f"知识点摘要_{stem}.docx"

            core.create_annotated_doc(text, kps, annotated_path)
            core.create_summary_doc(kps, original_filename, summary_path)

            jobs[job_id].update({
                "status": "done",
                "count": len(kps),
                "files": {
                    "annotated":      str(annotated_path),
                    "summary":        str(summary_path),
                    "annotated_name": annotated_path.name,
                    "summary_name":   summary_path.name,
                },
            })
        except Exception as e:
            jobs[job_id].update({"status": "error", "error": str(e)})
        finally:
            try:
                os.unlink(tmp_path)
            except OSError:
                pass

    threading.Thread(target=process, daemon=True).start()
    return redirect(url_for("result", job_id=job_id))


@app.route("/result/<job_id>")
def result(job_id):
    job = jobs.get(job_id, {"status": "error", "error": "任务不存在", "files": {}, "count": 0})
    return render_template_string(HTML, page="result", job=job, job_id=job_id)


@app.route("/status/<job_id>")
def status(job_id):
    job = jobs.get(job_id, {"status": "not_found"})
    return jsonify({"status": job["status"]})


@app.route("/download/<job_id>/<file_type>")
def download(job_id, file_type):
    job = jobs.get(job_id)
    if not job or job["status"] != "done":
        return "文件不存在", 404
    if file_type == "annotated":
        path = job["files"]["annotated"]
        name = job["files"]["annotated_name"]
    elif file_type == "summary":
        path = job["files"]["summary"]
        name = job["files"]["summary_name"]
    else:
        return "无效类型", 400
    return send_file(path, as_attachment=True, download_name=name)


# ─────────────────────────────────────────────
# 启动
# ─────────────────────────────────────────────

if __name__ == "__main__":
    if not core.API_KEY:
        print("错误：未找到 API_KEY，请检查 .env 文件。")
        sys.exit(1)

    import socket
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        local_ip = s.getsockname()[0]
        s.close()
    except Exception:
        local_ip = "127.0.0.1"

    print("\n" + "=" * 45)
    print("  智能笔记提取器  网页版已启动")
    print("=" * 45)
    print(f"  本机访问：  http://127.0.0.1:5000")
    print(f"  平板访问：  http://{local_ip}:5000")
    print("  （平板和电脑需在同一 Wi-Fi 网络）")
    print("=" * 45)
    print("  按 Ctrl+C 可停止服务\n")

    port = int(os.getenv("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
