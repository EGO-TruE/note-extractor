"""
智能笔记提取器
用法：python main.py <文件路径或文件夹路径>
"""

import os
import sys
import json
import re
import time
from pathlib import Path

import requests
from dotenv import load_dotenv

# ─────────────────────────────────────────────
# Section 1: 配置加载
# ─────────────────────────────────────────────

load_dotenv()

API_BASE_URL = os.getenv("API_BASE_URL", "https://www.traxnode.com/v1")
API_KEY      = os.getenv("API_KEY")
MODEL        = os.getenv("MODEL", "gpt-4o")
OUTPUT_DIR   = Path("output")


# ─────────────────────────────────────────────
# Section 2: 文件读取
# ─────────────────────────────────────────────

def read_pdf(path: Path) -> str:
    import pdfplumber
    pages = []
    with pdfplumber.open(path) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text()
            if text:
                pages.append(f"--- 第 {i + 1} 页 ---\n{text}")
    return "\n\n".join(pages)


def read_txt(path: Path) -> str:
    for encoding in ["utf-8", "utf-8-sig", "gbk"]:
        try:
            return path.read_text(encoding=encoding)
        except (UnicodeDecodeError, LookupError):
            continue
    raise ValueError(f"无法识别文件编码：{path.name}")


def read_markdown(path: Path) -> str:
    return read_txt(path)


def read_docx(path: Path) -> str:
    from docx import Document
    doc = Document(path)
    paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
    return "\n\n".join(paragraphs)


def read_pptx(path: Path) -> str:
    from pptx import Presentation
    prs = Presentation(path)
    slides = []
    for i, slide in enumerate(prs.slides):
        texts = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    line = para.text.strip()
                    if line:
                        texts.append(line)
        if texts:
            slides.append(f"--- 幻灯片 {i + 1} ---\n" + "\n".join(texts))
    return "\n\n".join(slides)


SUPPORTED_EXTENSIONS = {
    ".pdf":  read_pdf,
    ".txt":  read_txt,
    ".md":   read_markdown,
    ".docx": read_docx,
    ".pptx": read_pptx,
}


def read_file(path: Path) -> str:
    ext = path.suffix.lower()
    if ext not in SUPPORTED_EXTENSIONS:
        raise ValueError(f"不支持的文件类型：{ext}（文件：{path.name}）")
    text = SUPPORTED_EXTENSIONS[ext](path)
    if not text.strip():
        raise ValueError(f"文件内容为空或无法提取文本（可能是扫描版图片 PDF）：{path.name}")
    return text


def collect_files(input_path: Path) -> list:
    if input_path.is_file():
        return [input_path]
    elif input_path.is_dir():
        files = []
        for ext in SUPPORTED_EXTENSIONS:
            files.extend(input_path.glob(f"**/*{ext}"))
        return sorted(files)
    else:
        raise FileNotFoundError(f"路径不存在：{input_path}")


# ─────────────────────────────────────────────
# Section 3: AI 分析
# ─────────────────────────────────────────────

SYSTEM_PROMPT = """你是一位专业的知识点提取助手。请分析用户提供的文本，找出其中的重要知识点。

必须严格按照以下JSON格式输出，不要有任何额外文字：
{
  "知识点": [
    {
      "序号": 1,
      "标题": "知识点的简短标题（10字以内）",
      "原文片段": "该知识点在原文中对应的关键句子（必须精确匹配原文中的文字，用于定位标注）",
      "说明": "对该知识点的详细解释（50-150字）"
    }
  ]
}

提取规则：
1. 重点提取：概念定义、核心原理、重要结论、关键步骤、注意事项
2. 每个知识点必须有明确的学习价值
3. "原文片段"必须是原文中实际存在的文字，不能改写或添加
4. 输出纯JSON，不要有任何代码块标记（```）或额外说明
"""

USER_PROMPT_TEMPLATE = "请从以下文本中提取重要知识点：\n\n{text}"


def chunk_text(text: str, max_chars: int = 20000) -> list:
    paragraphs = text.split("\n\n")
    chunks = []
    current_chunk = []
    current_len = 0

    for para in paragraphs:
        para_len = len(para)
        if current_len + para_len > max_chars and current_chunk:
            chunks.append("\n\n".join(current_chunk))
            current_chunk = [para]
            current_len = para_len
        else:
            current_chunk.append(para)
            current_len += para_len

    if current_chunk:
        chunks.append("\n\n".join(current_chunk))

    return chunks


def call_ai(text_chunk: str, max_retries: int = 3) -> str:
    url = API_BASE_URL.rstrip("/") + "/chat/completions"
    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json",
    }
    payload = {
        "model": MODEL,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user",   "content": USER_PROMPT_TEMPLATE.format(text=text_chunk)},
        ],
        "temperature": 0.2,
    }
    last_error = None
    for attempt in range(max_retries):
        try:
            resp = requests.post(url, headers=headers, json=payload, timeout=300)
            resp.raise_for_status()
            return resp.json()["choices"][0]["message"]["content"]
        except Exception as e:
            last_error = e
            if attempt < max_retries - 1:
                print(f"  请求失败（第 {attempt + 1} 次），5 秒后重试：{e}")
                time.sleep(5)
    raise last_error


def parse_ai_response(response_text: str) -> list:
    # 清理可能的代码块标记
    cleaned = response_text.strip()
    cleaned = re.sub(r'^```(?:json)?\s*', '', cleaned)
    cleaned = re.sub(r'\s*```$', '', cleaned)

    try:
        data = json.loads(cleaned)
        return data.get("知识点", [])
    except json.JSONDecodeError:
        # 尝试从响应中提取 JSON 块
        match = re.search(r'\{[\s\S]*\}', cleaned)
        if match:
            try:
                data = json.loads(match.group())
                return data.get("知识点", [])
            except json.JSONDecodeError:
                pass
    print("  警告：AI 响应解析失败，跳过此块。")
    return []


def extract_knowledge_points(full_text: str, progress_callback=None) -> list:
    chunks = chunk_text(full_text)
    all_points = []
    total = len(chunks)
    print(f"  文本已分为 {total} 个块，正在逐块分析...")

    for i, chunk in enumerate(chunks):
        print(f"  正在分析第 {i + 1}/{total} 块...")
        if progress_callback:
            progress_callback(i + 1, total)
        try:
            response = call_ai(chunk)
            points = parse_ai_response(response)
            offset = len(all_points)
            for j, point in enumerate(points):
                point["序号"] = offset + j + 1
            all_points.extend(points)
            print(f"  第 {i + 1} 块提取到 {len(points)} 个知识点")
        except Exception as e:
            print(f"  警告：第 {i + 1} 块处理失败：{e}")
            continue

    return all_points


# ─────────────────────────────────────────────
# Section 4: 生成输出文档
# ─────────────────────────────────────────────

def create_annotated_doc(full_text: str, knowledge_points: list, output_path: Path):
    from docx import Document
    from docx.enum.text import WD_COLOR_INDEX

    doc = Document()
    doc.add_heading("原文（带知识点标注）", level=1)

    fragments = [kp.get("原文片段", "").strip() for kp in knowledge_points if kp.get("原文片段", "").strip()]

    for line in full_text.split("\n"):
        if not line.strip():
            doc.add_paragraph("")
            continue

        para = doc.add_paragraph()
        remaining = line

        # 找出本行中所有匹配的片段（取最早出现的那个）
        best_fragment = None
        best_idx = len(remaining) + 1

        for fragment in fragments:
            idx = remaining.find(fragment)
            if idx != -1 and idx < best_idx:
                best_idx = idx
                best_fragment = fragment

        if best_fragment is not None:
            # 片段前的普通文本
            if best_idx > 0:
                para.add_run(remaining[:best_idx])
            # 高亮片段
            run = para.add_run(best_fragment)
            run.bold = True
            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
            # 片段后的普通文本
            after = remaining[best_idx + len(best_fragment):]
            if after:
                para.add_run(after)
        else:
            para.add_run(remaining)

    doc.save(output_path)
    print(f"  已生成标注文档：{output_path.name}")


def create_summary_doc(knowledge_points: list, source_filename: str, output_path: Path):
    from docx import Document
    from docx.shared import RGBColor

    doc = Document()
    doc.add_heading("知识点摘要", level=1)

    subtitle = doc.add_paragraph()
    run = subtitle.add_run(f"来源文件：{source_filename}")
    run.italic = True
    run.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

    doc.add_paragraph()

    if not knowledge_points:
        doc.add_paragraph("未提取到知识点。")
        doc.save(output_path)
        return

    for kp in knowledge_points:
        seq   = kp.get("序号", "?")
        title = kp.get("标题", "（无标题）")
        frag  = kp.get("原文片段", "").strip()
        desc  = kp.get("说明", "").strip()

        doc.add_heading(f"{seq}. {title}", level=2)

        if frag:
            quote = doc.add_paragraph(style="Quote")
            quote.add_run(f"原文：{frag}").italic = True

        if desc:
            doc.add_paragraph(desc)

        doc.add_paragraph()

    doc.save(output_path)
    print(f"  已生成摘要文档：{output_path.name}")


# ─────────────────────────────────────────────
# Section 5: 主入口
# ─────────────────────────────────────────────

def main():
    if not API_KEY:
        print("错误：未找到 API_KEY，请将 .env.example 复制为 .env 并填写 API_KEY。")
        sys.exit(1)

    if len(sys.argv) < 2:
        print("用法：python main.py <文件路径或文件夹路径>")
        print()
        print("支持的文件格式：PDF、TXT、Markdown (.md)、Word (.docx)、PPT (.pptx)")
        print("示例：")
        print("  python main.py notes.pdf")
        print("  python main.py ./lecture_slides/")
        sys.exit(1)

    input_path = Path(sys.argv[1])
    OUTPUT_DIR.mkdir(exist_ok=True)

    try:
        files = collect_files(input_path)
    except FileNotFoundError as e:
        print(f"错误：{e}")
        sys.exit(1)

    if not files:
        print("警告：未找到支持的文件。")
        sys.exit(0)

    print(f"共找到 {len(files)} 个文件待处理。\n")

    for file_path in files:
        print(f"正在处理：{file_path.name}")
        print("-" * 50)

        try:
            full_text = read_file(file_path)
        except Exception as e:
            print(f"  跳过：{e}\n")
            continue

        print(f"  文本提取完成，共 {len(full_text)} 个字符")

        knowledge_points = extract_knowledge_points(full_text)

        if not knowledge_points:
            print("  警告：未提取到任何知识点，跳过生成文档。\n")
            continue

        print(f"  共提取到 {len(knowledge_points)} 个知识点，正在生成文档...")

        stem = file_path.stem
        annotated_path = OUTPUT_DIR / f"原文_带标注_{stem}.docx"
        summary_path   = OUTPUT_DIR / f"知识点摘要_{stem}.docx"

        create_annotated_doc(full_text, knowledge_points, annotated_path)
        create_summary_doc(knowledge_points, file_path.name, summary_path)

        print(f"  处理完成！\n")

    print(f"全部处理完毕。输出文件位于：{OUTPUT_DIR.resolve()}")


if __name__ == "__main__":
    main()
