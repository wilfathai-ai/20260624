import os
import json
import re
import subprocess
import tempfile
import requests
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import fitz  # PyMuPDF
from docx import Document

app = Flask(__name__)
CORS(app)

OUTPUT_DIR = os.path.join(os.path.dirname(__file__), '..', 'output')
os.makedirs(OUTPUT_DIR, exist_ok=True)

OLLAMA_URL = "http://localhost:11434"

# ─── カラーパレット定義 ─────────────────────────────────────────────────
PALETTES = {
    "ブルー系":         {"bg": "1E3A5F", "accent": "4A90D9", "text": "FFFFFF", "sub": "BDD5EA"},
    "グリーン系":       {"bg": "1B4332", "accent": "52B788", "text": "FFFFFF", "sub": "B7E4C7"},
    "レッド系":         {"bg": "7B1D1D", "accent": "E63946", "text": "FFFFFF", "sub": "FFCCD5"},
    "モノクロ":         {"bg": "212121", "accent": "BDBDBD", "text": "FFFFFF", "sub": "E0E0E0"},
    "ウォームオレンジ": {"bg": "7C2D12", "accent": "EA580C", "text": "FFFFFF", "sub": "FED7AA"},
    "パープル系":       {"bg": "3B1F6E", "accent": "9B5DE5", "text": "FFFFFF", "sub": "D4B8FA"},
}

# ─── Ollama ヘルパー ───────────────────────────────────────────────────
def get_available_models():
    try:
        r = requests.get(f"{OLLAMA_URL}/api/tags", timeout=5)
        if r.ok:
            return [m["name"] for m in r.json().get("models", [])]
    except Exception:
        pass
    return []

def ollama_generate(model: str, prompt: str) -> str:
    payload = {"model": model, "prompt": prompt, "stream": False}
    r = requests.post(f"{OLLAMA_URL}/api/generate", json=payload, timeout=120)
    r.raise_for_status()
    return r.json().get("response", "")

# ─── ファイル解析 ──────────────────────────────────────────────────────
def extract_text(path: str) -> str:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        doc = fitz.open(path)
        return "\n".join(page.get_text() for page in doc)
    if ext in (".docx", ".doc"):
        doc = Document(path)
        return "\n".join(p.text for p in doc.paragraphs if p.text.strip())
    with open(path, encoding="utf-8", errors="ignore") as f:
        return f.read()

# ─── プロンプト生成 ────────────────────────────────────────────────────
def build_prompt(outline_text: str, design: dict, slide_count: int) -> str:
    return f"""あなたはプロのプレゼンテーションデザイナーです。
以下のアウトラインをもとに、{slide_count}枚のスライドデータをJSON形式で生成してください。

## デザイン設定
- テーマ: {design.get('theme', 'プロフェッショナル')}
- カラー: {design.get('colorScheme', 'ブルー系')}
- フォントスタイル: {design.get('fontStyle', 'モダン')}
- レイアウト: {design.get('layout', 'タイトル＋コンテンツ')}

## アウトライン
{outline_text}

## 出力形式（JSONのみ・説明文不要・コードブロック不要）
{{
  "title": "プレゼンタイトル",
  "theme": "{design.get('theme', '')}",
  "colorScheme": "{design.get('colorScheme', 'ブルー系')}",
  "slides": [
    {{
      "slideNumber": 1,
      "type": "title",
      "title": "スライドタイトル",
      "subtitle": "サブタイトル（任意）",
      "content": [],
      "speakerNotes": "発表者メモ"
    }},
    {{
      "slideNumber": 2,
      "type": "content",
      "title": "スライドタイトル",
      "subtitle": "",
      "content": ["ポイント1", "ポイント2", "ポイント3"],
      "speakerNotes": "発表者メモ"
    }}
  ]
}}"""

# ─── PPTX生成（Node.js / pptxgenjs） ─────────────────────────────────
def build_pptx(slides_data: dict, output_path: str) -> str:
    palette = PALETTES.get(slides_data.get("colorScheme", "ブルー系"), PALETTES["ブルー系"])

    js = f"""
const PptxGenJS = require('pptxgenjs');
const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_16x9';
pptx.title = {json.dumps(slides_data.get('title', 'プレゼンテーション'))};

const BG  = '#{palette["bg"]}';
const ACC = '#{palette["accent"]}';
const TXT = '#{palette["text"]}';
const SUB = '#{palette["sub"]}';
const slides = {json.dumps(slides_data.get('slides', []))};

slides.forEach((s) => {{
  const slide = pptx.addSlide();
  slide.background = {{ color: BG }};

  if (s.type === 'title') {{
    slide.addText(s.title || '', {{
      x:0.5, y:2.5, w:9, h:1.4,
      fontSize:40, bold:true, color:TXT, align:'center', fontFace:'Calibri'
    }});
    if (s.subtitle) slide.addText(s.subtitle, {{
      x:0.5, y:4.0, w:9, h:0.8,
      fontSize:20, color:SUB, align:'center', fontFace:'Calibri'
    }});
    slide.addShape(pptx.ShapeType.rect, {{ x:2.5, y:5.1, w:5, h:0.08, fill:{{ color:ACC }} }});
  }} else {{
    slide.addShape(pptx.ShapeType.rect, {{ x:0, y:0, w:10, h:1.1, fill:{{ color:ACC }} }});
    slide.addText(s.title || '', {{
      x:0.4, y:0.15, w:9.2, h:0.8,
      fontSize:26, bold:true, color:TXT, fontFace:'Calibri'
    }});
    const items = (s.content || []).filter(c => c && c.trim());
    if (items.length) {{
      const bullets = items.map(c => ({{
        text: c,
        options: {{ bullet:{{type:'bullet'}}, fontSize:17, color:TXT, breakLine:true, paraSpaceAfter:8 }}
      }}));
      slide.addText(bullets, {{ x:0.6, y:1.4, w:8.8, h:4.8, fontFace:'Calibri', valign:'top' }});
    }}
    if (s.subtitle) slide.addText(s.subtitle, {{
      x:0.6, y:6.5, w:8.8, h:0.4,
      fontSize:12, color:SUB, italic:true, fontFace:'Calibri'
    }});
  }}
  if (s.speakerNotes) slide.addNotes(s.speakerNotes);
}});

pptx.writeFile({{ fileName: {json.dumps(output_path)} }})
  .then(() => console.log('OK'))
  .catch(e => {{ console.error(e); process.exit(1); }});
"""
    with tempfile.NamedTemporaryFile(suffix=".js", delete=False, mode="w") as f:
        f.write(js); tmp = f.name
    try:
        env = os.environ.copy()
        global_modules = subprocess.run(
            ["npm", "root", "-g"], capture_output=True, text=True
        ).stdout.strip()
        if global_modules:
            env["NODE_PATH"] = global_modules
        r = subprocess.run(["node", tmp], capture_output=True, text=True, timeout=60, env=env)
        if r.returncode != 0:
            raise RuntimeError(r.stderr)
    finally:
        os.unlink(tmp)
    return output_path

# ─── Google Slides用JSON生成 ──────────────────────────────────────────
def build_google_slides_json(slides_data: dict, output_path: str) -> str:
    palette = PALETTES.get(slides_data.get("colorScheme", "ブルー系"), PALETTES["ブルー系"])
    export = {
        "presentationTitle": slides_data.get("title", "プレゼンテーション"),
        "colorScheme": palette,
        "slides": slides_data.get("slides", []),
    }
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(export, f, ensure_ascii=False, indent=2)
    return output_path

# ─── JSONパース＆展開処理 ──────────────────────────────────────────────
def parse_slides_json(raw: str) -> dict:
    cleaned = re.sub(r"```(?:json)?|```", "", raw).strip()
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        m = re.search(r'\{[\s\S]*\}', cleaned)
        if m:
            return json.loads(m.group())
        raise ValueError("JSONのパースに失敗しました")

# ─── APIエンドポイント ─────────────────────────────────────────────────
@app.route("/health")
def health():
    return jsonify({"status": "ok"})

@app.route("/api/models")
def api_models():
    return jsonify({"models": get_available_models()})

@app.route("/api/preview", methods=["POST"])
def api_preview():
    """プレビュー用（スライドJSONを返す、ファイルは生成しない）"""
    model       = request.form.get("model", "")
    design      = json.loads(request.form.get("design", "{}"))
    slide_count = int(request.form.get("slideCount", 8))
    file        = request.files.get("file")

    if not model or not file:
        return jsonify({"error": "モデルとファイルが必要です"}), 400

    suffix = os.path.splitext(file.filename)[1].lower() or ".txt"
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        file.save(tmp.name); tmp_path = tmp.name

    try:
        text = extract_text(tmp_path)
    finally:
        os.unlink(tmp_path)

    if not text.strip():
        return jsonify({"error": "テキストを読み取れませんでした"}), 400

    try:
        raw = ollama_generate(model, build_prompt(text[:4000], design, slide_count))
        data = parse_slides_json(raw)
        data["colorScheme"] = design.get("colorScheme", "ブルー系")
        return jsonify(data)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/generate", methods=["POST"])
def api_generate():
    """ファイル生成してダウンロード"""
    model       = request.form.get("model", "")
    design      = json.loads(request.form.get("design", "{}"))
    slide_count = int(request.form.get("slideCount", 8))
    export_fmt  = request.form.get("exportFormat", "pptx")
    file        = request.files.get("file")

    if not model or not file:
        return jsonify({"error": "モデルとファイルが必要です"}), 400

    suffix = os.path.splitext(file.filename)[1].lower() or ".txt"
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        file.save(tmp.name); tmp_path = tmp.name

    try:
        text = extract_text(tmp_path)
    finally:
        os.unlink(tmp_path)

    try:
        raw = ollama_generate(model, build_prompt(text[:4000], design, slide_count))
        data = parse_slides_json(raw)
        data["colorScheme"] = design.get("colorScheme", "ブルー系")
    except Exception as e:
        return jsonify({"error": str(e)}), 500

    if export_fmt == "gslides":
        out = os.path.join(OUTPUT_DIR, "slides_output.json")
        build_google_slides_json(data, out)
        return send_file(out, as_attachment=True, download_name="google_slides_data.json",
                         mimetype="application/json")
    else:
        out = os.path.join(OUTPUT_DIR, "slides_output.pptx")
        try:
            build_pptx(data, out)
        except Exception as e:
            return jsonify({"error": f"PPTX生成エラー: {e}"}), 500
        return send_file(out, as_attachment=True, download_name="presentation.pptx",
                         mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation")

if __name__ == "__main__":
    app.run(debug=True, port=5050)
