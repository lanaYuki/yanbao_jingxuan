"""
ai_helper.py - 调用 DeepSeek API 实现 AI 辅助功能
使用 openai 兼容接口。

安全改进：优先从环境变量 DEEPSEEK_API_KEY 读取 API Key，避免将密钥硬编码到源码中。
如果未提供环境变量，仍然会使用仓库中原有的回退值（请尽快删除回退值并在运行环境中设置变量）。
"""

import os
import sys
from openai import OpenAI

# 从环境变量读取 API Key
_api_key = os.environ.get('DEEPSEEK_API_KEY')
if not _api_key:
    sys.exit(
        '错误：未设置 DEEPSEEK_API_KEY 环境变量。\n'
        '请在运行前执行：export DEEPSEEK_API_KEY=your_key_here'
    )

_client = OpenAI(
    api_key=_api_key,
    base_url='https://api.deepseek.com/v1',
)

KNOWN_SECTIONS = ['宏观专栏', '专题聚焦', '策略聚焦', '固定收益', '行业聚焦', '港股策略', '中国策略']


def suggest_section(title: str, abstract_text: str) -> str:
    """
    根据研报标题和摘要文字，推荐最合适的栏目名。
    返回栏目名字符串。
    """
    sections_str = '、'.join(KNOWN_SECTIONS)
    prompt = (
        f'你是一名研报编辑，请根据以下研报的标题和摘要，从备选栏目中选择最合适的一个栏目名。\n'
        f'备选栏目：{sections_str}\n'
        f'研报标题：{title}\n'
        f'研报摘要：{abstract_text[:300]}\n\n'
        f'只返回栏目名本身，不要任何解释。'
    )
    try:
        resp = _client.chat.completions.create(
            model='deepseek-chat',
            messages=[{'role': 'user', 'content': prompt}],
            max_tokens=20,
            temperature=0,
        )
        result = resp.choices[0].message.content.strip()
        # 验证返回值在已知栏目中
        for s in KNOWN_SECTIONS:
            if s in result:
                return s
        return result
    except Exception as e:
        return ''



def check_and_fix_orphan_lines(
    paragraphs: list,
    page_lines: float,
    chars_per_line: int = 33,
) -> dict:
    """
    检查高亮段落列表在排版后是否会产生孤行，并给出修复建议。

    paragraphs: [{'para_text': str, 'runs': [...]}]
    page_lines: 每页可用行数（已减去标题、归属信息等占用）
    chars_per_line: 每行字符数

    返回：
    {
      'has_orphan': bool,
      'line_spacing': float,   # 建议行距（23~24之间）
      'paragraphs': [{'para_text': str, 'runs': [...]}],  # 可能合并/删减后的段落
      'reason': str,
    }
    """
    import math

    def count_lines(text):
        return math.ceil((len(text) + 2) / chars_per_line)

    total_lines = sum(count_lines(p['para_text']) for p in paragraphs)

    # 判断是否有孤行：如果总行数模page_lines余1（最后一页只有1行）
    last_page_lines = total_lines % page_lines
    has_orphan = (0 < last_page_lines <= 2)

    if not has_orphan:
        return {
            'has_orphan': False,
            'line_spacing': 24,
            'paragraphs': paragraphs,
            'reason': '无孤行',
        }

    # 让AI决定如何处理
    para_texts = [p['para_text'] for p in paragraphs]
    prompt = (
        f'你是一名研报排版编辑。以下研报摘录段落在排版后（每页{page_lines:.0f}行，每行约{chars_per_line}字）'
        f'会在最后一页产生孤行（只有{last_page_lines:.0f}行内容），需要调整。\n\n'
        f'段落内容：\n'
        + '\n---\n'.join(f'[{i+1}] {t}' for i, t in enumerate(para_texts))
        + f'\n\n请选择以下方案之一解决孤行问题：\n'
        f'A. 建议行距：给出23~24之间的行距值（如23.5），使总行数减少1-2行\n'
        f'B. 合并相邻段落：指定哪两段合并（如"合并段落1和2"），并给出合并后的文字\n'
        f'C. 删减某段末尾：指定哪段删减几个字，并给出删减后的文字\n\n'
        f'注意：优先选A（行距微调），仅当A无法解决时才考虑B或C。\n'
        f'返回JSON格式：\n'
        f'{{"method": "A", "line_spacing": 23.5}}\n'
        f'或 {{"method": "B", "merge": [1, 2], "merged_text": "合并后的文字"}}\n'
        f'或 {{"method": "C", "para_index": 3, "new_text": "删减后的文字"}}\n'
        f'只返回JSON，不要任何解释。'
    )

    try:
        resp = _client.chat.completions.create(
            model='deepseek-chat',
            messages=[{'role': 'user', 'content': prompt}],
            max_tokens=300,
            temperature=0,
        )
        import json, re
        raw = resp.choices[0].message.content.strip()
        # 提取JSON
        m = re.search(r'\{.*\}', raw, re.DOTALL)
        if not m:
            return {'has_orphan': True, 'line_spacing': 23.5,
                    'paragraphs': paragraphs, 'reason': 'AI返回格式错误，默认调整行距'}
        result = json.loads(m.group(0))
        method = result.get('method', 'A')

        if method == 'A':
            ls = float(result.get('line_spacing', 23.5))
            ls = max(23.0, min(24.0, ls))
            return {'has_orphan': True, 'line_spacing': ls,
                    'paragraphs': paragraphs, 'reason': f'AI建议行距{ls}'}

        elif method == 'B':
            idxs = result.get('merge', [])
            merged_text = result.get('merged_text', '')
            if len(idxs) == 2 and merged_text:
                i1 = idxs[0] - 1
                new_paras = []
                skip = False
                for idx, p in enumerate(paragraphs):
                    if skip:
                        skip = False
                        continue
                    if idx == i1:
                        merged = dict(p)
                        merged['para_text'] = merged_text
                        merged['runs'] = [{'text': merged_text, 'bold': False}]
                        new_paras.append(merged)
                        skip = True
                    else:
                        new_paras.append(p)
                return {'has_orphan': True, 'line_spacing': 24,
                        'paragraphs': new_paras, 'reason': f'AI建议合并段落{idxs}'}

        elif method == 'C':
            idx = result.get('para_index', 1) - 1
            new_text = result.get('new_text', '')
            if 0 <= idx < len(paragraphs) and new_text:
                new_paras = list(paragraphs)
                new_paras[idx] = dict(paragraphs[idx])
                new_paras[idx]['para_text'] = new_text
                new_paras[idx]['runs'] = [{'text': new_text, 'bold': False}]
                return {'has_orphan': True, 'line_spacing': 24,
                        'paragraphs': new_paras, 'reason': f'AI建议删减段落{idx+1}'}

        return {'has_orphan': True, 'line_spacing': 23.5,
                'paragraphs': paragraphs, 'reason': 'AI方案无法解析，默认调整行距'}

    except Exception as e:
        return {'has_orphan': True, 'line_spacing': 23.5,
                'paragraphs': paragraphs, 'reason': f'AI调用失败：{e}'}


def generate_summary(title: str, highlighted_paragraphs: list, max_lines: int = 3,
                     is_first: bool = False) -> str:
    """
    将研报高亮内容压缩为摘要页导读。
    is_first=True：第一篇，目标90-120字（对应4行，每行约20字）。
    is_first=False：其余篇，目标60-90字（对应3行）。
    不换行，输出一整段文字。
    将"我们认为"统一改为"报告认为"。
    """
    if is_first:
        char_min, char_max = 90, 120
        line_desc = '4行（约90-120字）'
    else:
        char_min, char_max = 60, 90
        line_desc = '3行（约60-90字）'

    full_text = '\n'.join(p['para_text'] for p in highlighted_paragraphs)
    prompt = (
        f'你是一名研报编辑，请将以下研报内容压缩为封面摘要导读。\n'
        f'要求：\n'
        f'1. 总字数控制在{char_min}到{char_max}字之间（目标{line_desc}，每行约30字）\n'
        f'2. 不要换行，输出一整段连续文字\n'
        f'3. 将所有"我们认为"改为"报告认为"\n'
        f'4. 只保留最核心的观点和结论\n'
        f'5. 直接输出摘要内容，不要标题、序号或任何解释\n\n'
        f'研报标题：{title}\n'
        f'研报内容：\n{full_text[:1500]}'
    )
    try:
        resp = _client.chat.completions.create(
            model='deepseek-chat',
            messages=[{'role': 'user', 'content': prompt}],
            max_tokens=150,
            temperature=0.3,
        )
        result = resp.choices[0].message.content.strip()
        # 移除任何换行，确保是一段
        result = result.replace('\n', ' ').replace('\r', '')
        return result
    except Exception as e:
        return ''
