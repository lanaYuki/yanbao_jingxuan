"""
app.py - 研报精选自动化工具 Flask 后端
"""

import os
import uuid
import threading
from flask import Flask, request, jsonify, send_file, render_template
from werkzeug.utils import secure_filename

from extractor import extract_all
from builder import build_jingxuan
from ai_helper import suggest_section, generate_summary
from list_builder import parse_pasted_text, build_list_docx, CATEGORIES

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB

UPLOAD_DIR = os.path.join(os.path.dirname(__file__), 'uploads')
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), 'output')
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# 存储每个上传任务的 AI 分析结果（临时内存缓存）
_ai_cache: dict = {}
_ai_lock = threading.Lock()


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/analyze', methods=['POST'])
def analyze():
    """
    上传单个研报 docx，返回提取信息 + 触发 AI 分析。
    前端上传4个文件时分4次调用，或一次性上传。
    """
    file = request.files.get('file')
    if not file or not file.filename.endswith('.docx'):
        return jsonify({'error': '请上传 .docx 文件'}), 400

    # 保存文件
    file_id = str(uuid.uuid4())
    filename = secure_filename(file.filename)
    save_path = os.path.join(UPLOAD_DIR, f'{file_id}_{filename}')
    file.save(save_path)

    # 提取基础信息
    try:
        data = extract_all(save_path)
    except Exception as e:
        return jsonify({'error': f'解析失败：{str(e)}'}), 500

    result = {
        'file_id': file_id,
        'file_path': save_path,
        'title': data['title'],
        'date': data['date'],
        'institution': data['institution'],
        'authors': data['authors'],
        'highlighted_count': len(data['highlighted_paragraphs']),
        'ai_section': None,    # AI推荐栏目（异步填充）
        'ai_summary': None,    # AI摘要（异步填充）
        'ai_ready': False,
    }

    # 缓存提取结果供后续生成使用
    with _ai_lock:
        _ai_cache[file_id] = {
            'save_path': save_path,
            'data': data,
            'ai_section': None,
            'ai_summary': None,
        }

    # 异步调用 AI（is_first 由生成时根据顺序决定，这里先用非第一篇默认值生成）
    def run_ai():
        ai_section = suggest_section(data['title'], data['abstract_text'])
        ai_summary = generate_summary(data['title'], data['highlighted_paragraphs'], is_first=False)
        with _ai_lock:
            if file_id in _ai_cache:
                _ai_cache[file_id]['ai_section'] = ai_section
                _ai_cache[file_id]['ai_summary'] = ai_summary

    threading.Thread(target=run_ai, daemon=True).start()

    return jsonify(result)


@app.route('/ai_result/<file_id>')
def ai_result(file_id):
    """轮询获取 AI 分析结果。"""
    with _ai_lock:
        cache = _ai_cache.get(file_id)
    if not cache:
        return jsonify({'error': 'not found'}), 404

    ready = cache['ai_section'] is not None
    return jsonify({
        'ready': ready,
        'ai_section': cache.get('ai_section'),
        'ai_summary': cache.get('ai_summary'),
    })


@app.route('/generate', methods=['POST'])
def generate():
    """
    接收4篇研报的配置，生成精选 docx 并返回下载链接。

    请求体（JSON）：
    {
      "issue": "259",
      "articles": [
        {
          "file_id": "...",
          "section": "专题聚焦",   // 用户确认的栏目名
        },
        ...
      ]
    }
    """
    body = request.get_json()
    issue = body.get('issue', '')
    articles_config = body.get('articles', [])
    # link_map: [{name: str, url: str}, ...] -> {name: url}
    link_map_list = body.get('link_map', [])
    link_map = {item['name']: item['url'] for item in link_map_list if item.get('name') and item.get('url')}

    if not articles_config:
        return jsonify({'error': '请提供文章配置'}), 400

    articles = []
    for idx, cfg in enumerate(articles_config):
        file_id = cfg.get('file_id')
        section = cfg.get('section', '')
        summary = cfg.get('summary', '')   # 用户编辑后的摘要

        with _ai_lock:
            cache = _ai_cache.get(file_id)
        if not cache:
            return jsonify({'error': f'文件 {file_id} 未找到，请重新上传'}), 400

        data = cache['data']

        # 如果没有摘要，按顺序重新生成（第一篇60-80字，其余40-60字）
        if not summary:
            is_first = (idx == 0)
            summary = generate_summary(data['title'], data['highlighted_paragraphs'], is_first=is_first)

        articles.append({
            'section': section,
            'title': data['title'],
            'paragraphs': data['highlighted_paragraphs'],
            'date': data['date'],
            'institution': data['institution'],
            'authors': data['authors'],
            'footnotes': data['footnotes'],
            'summary': summary,
        })

    # 生成输出文件名
    issue_str = f'第{issue}期' if issue else ''
    output_filename = f'CGI每周研报精选（{issue_str}）-摘要.docx'
    output_path = os.path.join(OUTPUT_DIR, output_filename)

    try:
        build_jingxuan(articles, output_path, issue=issue, link_map=link_map)
    except Exception as e:
        return jsonify({'error': f'生成失败：{str(e)}'}), 500

    return jsonify({
        'success': True,
        'download_url': f'/download/{output_filename}',
        'filename': output_filename,
    })


@app.route('/download/<filename>')
def download(filename):
    """下载生成的精选 docx。"""
    file_path = os.path.join(OUTPUT_DIR, filename)
    if not os.path.exists(file_path):
        return jsonify({'error': '文件不存在'}), 404
    return send_file(file_path, as_attachment=True, download_name=filename)


@app.route('/parse_list', methods=['POST'])
def parse_list():
    """
    解析各分类粘贴文本，提取研报标题和日期。
    请求体：{ "categories": {"宏观": "...", "策略及大宗商品": "...", ...} }
    返回：{ "宏观": [{"title": ..., "date": ...}, ...], ... }
    """
    body = request.get_json()
    categories = body.get('categories', {})
    result = {}
    for cat in CATEGORIES:
        text = categories.get(cat, '').strip()
        if text:
            result[cat] = parse_pasted_text(text)
        else:
            result[cat] = []
    return jsonify(result)


@app.route('/generate_list', methods=['POST'])
def generate_list():
    """
    生成研报清单 docx。
    请求体：{
      "issue": "259",
      "categories": {
        "宏观": [{"title": str, "date": str}, ...],
        ...
      }
    }
    """
    body = request.get_json()
    issue = body.get('issue', '')
    category_data = body.get('categories', {})

    issue_str = f'第{issue}期' if issue else ''
    output_filename = f'CGI每周研报精选（{issue_str}）-研报清单.docx'
    output_path = os.path.join(OUTPUT_DIR, output_filename)

    try:
        build_list_docx(category_data, output_path, issue=issue)
    except Exception as e:
        return jsonify({'error': f'生成失败：{str(e)}'}), 500

    return jsonify({
        'success': True,
        'download_url': f'/download/{output_filename}',
        'filename': output_filename,
    })


if __name__ == '__main__':
    app.run(debug=True, port=5000)
