# -*- coding: utf-8 -*-
#!/usr/bin/env python3
"""
CXskill - 程序文件生成脚本（docx 处理库，AI 驱动）
=====================================================
与 SCskill/generate_manual.py 结构一致，用于生成程序文件。
不同点：
  - find_template 查找"程序文件"分类而非"手册"
  - 全质知识库路径：external_kb/体系文件/程序文件/{部门}/
  - 企业内部文件路径：agent_{id}/程序文件/{部门}/
"""
import os
import sys
import json
import re
import subprocess
import tempfile
from pathlib import Path
from datetime import datetime


def ensure_packages():
    import importlib
    try:
        importlib.import_module('docx')
    except ImportError:
        subprocess.check_call([sys.executable, '-m', 'pip', 'install',
                               'python-docx', '--quiet'])


ensure_packages()

from docx import Document

SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_ROOT = SCRIPT_DIR.parent
TEMPLATES_DIR = SKILL_ROOT / "templates"


# ===================================================================
# 1. 模板查找（三级：企业内部→全质知识库→无内置兜底）
# ===================================================================

def find_template(agent_id=None, documents_dir=None):
    """查找程序文件模板（三级查找）：
    1. 企业内部文件知识库（agent_{agent_id}/程序文件/）递归搜索所有子目录
    2. 全质知识库（external_kb/体系文件/程序文件/）递归搜索所有子目录
    3. 无内置模板（程序文件无通用模板，必须从知识库找）

    返回 (template_path, need_convert, template_source, template_subdir)
    template_source: 'internal' / 'external' / None
    template_subdir: 模板所在的部门子目录名（如"总经办二级"）
    """
    if documents_dir is None:
        candidate = SKILL_ROOT.parent.parent / "data" / "documents"
        if candidate.exists():
            documents_dir = candidate
        else:
            documents_dir = SKILL_ROOT.parent / "data" / "documents"
    else:
        documents_dir = Path(documents_dir)

    print(f"[INFO] find_template: documents_dir={documents_dir}, agent_id={agent_id}")

    # 1. 企业内部文件知识库
    if agent_id:
        proc_dir = documents_dir / f"agent_{agent_id}" / "程序文件"
        print(f"[INFO] 查找企业内部文件: {proc_dir} (exists={proc_dir.exists()})")
        if proc_dir.exists():
            for root, dirs, files in os.walk(str(proc_dir)):
                for f in sorted(files):
                    if f.lower().endswith('.docx') and not f.startswith('~$'):
                        rel = os.path.relpath(root, str(proc_dir))
                        print(f"[INFO] 从企业内部文件知识库找到模板: {f} (路径: {root})")
                        return Path(root) / f, False, 'internal', rel
            for root, dirs, files in os.walk(str(proc_dir)):
                for f in sorted(files):
                    if f.lower().endswith('.doc') and not f.startswith('~$'):
                        rel = os.path.relpath(root, str(proc_dir))
                        print(f"[INFO] 从企业内部文件知识库找到 .doc 模板: {f} (路径: {root})")
                        return Path(root) / f, True, 'internal', rel
            print(f"[INFO] 企业内部文件知识库 程序文件/ 下未找到文件")

    # 2. 全质知识库
    ext_proc_dir = documents_dir / "external_kb" / "体系文件" / "程序文件"
    print(f"[INFO] 查找全质知识库: {ext_proc_dir} (exists={ext_proc_dir.exists()})")
    if ext_proc_dir.exists():
        for root, dirs, files in os.walk(str(ext_proc_dir)):
            for f in sorted(files):
                if f.lower().endswith('.docx') and not f.startswith('~$'):
                    rel = os.path.relpath(root, str(ext_proc_dir))
                    print(f"[INFO] 从全质知识库找到模板: {f} (路径: {root})")
                    return Path(root) / f, False, 'external', rel
        for root, dirs, files in os.walk(str(ext_proc_dir)):
            for f in sorted(files):
                if f.lower().endswith('.doc') and not f.startswith('~$'):
                    rel = os.path.relpath(root, str(ext_proc_dir))
                    print(f"[INFO] 从全质知识库找到 .doc 模板: {f} (路径: {root})")
                    return Path(root) / f, True, 'external', rel

    # 3. 无内置模板
    print(f"[INFO] 未找到任何程序文件模板")
    return None, False, None, None


def find_all_templates(agent_id=None, documents_dir=None):
    """查找所有程序文件模板（收集所有部门的所有文件）

    查找逻辑：
    1. 企业内部文件 agent_{id}/程序文件/ → 递归搜索所有子目录和文件
    2. 全质知识库 external_kb/体系文件/程序文件/ → 递归搜索所有子目录和文件
    3. 按部门去重：如果企业内部文件有某部门的模板，全质知识库的同部门模板跳过

    返回 list of dict:
    [
      {"path": Path, "need_convert": bool, "source": "internal"/"external",
       "dept": "部门名", "filename": "文件名"},
      ...
    ]
    """
    if documents_dir is None:
        candidate = SKILL_ROOT.parent.parent / "data" / "documents"
        if candidate.exists():
            documents_dir = candidate
        else:
            documents_dir = SKILL_ROOT.parent / "data" / "documents"
    else:
        documents_dir = Path(documents_dir)

    print(f"[INFO] find_all_templates: documents_dir={documents_dir}, agent_id={agent_id}")

    # 收集企业内部文件
    internal_templates = []
    if agent_id:
        proc_dir = documents_dir / f"agent_{agent_id}" / "程序文件"
        print(f"[INFO] 查找企业内部文件: {proc_dir} (exists={proc_dir.exists()})")
        if proc_dir.exists():
            for root, dirs, files in os.walk(str(proc_dir)):
                for f in sorted(files):
                    if f.startswith('~$'):
                        continue
                    ext = f.lower().rsplit('.', 1)[-1] if '.' in f else ''
                    if ext in ('docx', 'doc'):
                        rel = os.path.relpath(root, str(proc_dir))
                        # 部门名取第一级子目录（如"质量部二级"），如果文件直接在根目录则 dept="通用"
                        parts = rel.replace('\\', '/').split('/')
                        dept = parts[0] if parts and parts[0] != '.' else '通用'
                        internal_templates.append({
                            "path": Path(root) / f,
                            "need_convert": (ext == 'doc'),
                            "source": 'internal',
                            "dept": dept,
                            "filename": f
                        })
                        print(f"[INFO] 企业内部文件: {dept}/{f} (source=internal)")

    # 收集全质知识库
    ext_templates = []
    ext_proc_dir = documents_dir / "external_kb" / "体系文件" / "程序文件"
    print(f"[INFO] 查找全质知识库: {ext_proc_dir} (exists={ext_proc_dir.exists()})")
    if ext_proc_dir.exists():
        for root, dirs, files in os.walk(str(ext_proc_dir)):
            for f in sorted(files):
                if f.startswith('~$'):
                    continue
                ext = f.lower().rsplit('.', 1)[-1] if '.' in f else ''
                if ext in ('docx', 'doc'):
                    rel = os.path.relpath(root, str(ext_proc_dir))
                    parts = rel.replace('\\', '/').split('/')
                    dept = parts[0] if parts and parts[0] != '.' else '通用'
                    ext_templates.append({
                        "path": Path(root) / f,
                        "need_convert": (ext == 'doc'),
                        "source": 'external',
                        "dept": dept,
                        "filename": f
                    })
                    print(f"[INFO] 全质知识库: {dept}/{f} (source=external)")

    # 不再按部门整体去重，改为返回所有模板
    # 同部门去重逻辑交给 API 层用 AI 智能判断（按文件内容主题匹配，而非按部门整体跳过）
    all_templates = internal_templates + ext_templates
    print(f"[INFO] 总计找到 {len(all_templates)} 个模板（企业内部 {len(internal_templates)} + 全质知识库 {len(ext_templates)}）")
    for t in all_templates:
        print(f"  - [{t['source']}] {t['dept']}/{t['filename']}")

    return all_templates


def convert_doc_to_docx(doc_path):
    """用 LibreOffice 把 .doc 转成 .docx"""
    tmp_dir = tempfile.mkdtemp(prefix='doc2docx_')
    soffice_paths = [
        'soffice', '/usr/bin/soffice',
        'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
        'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe',
    ]
    for sp in soffice_paths:
        try:
            result = subprocess.run(
                [sp, '--headless', '--convert-to', 'docx',
                 str(doc_path), '--outdir', tmp_dir],
                capture_output=True, text=True, timeout=120
            )
            if result.returncode == 0:
                basename = os.path.splitext(os.path.basename(str(doc_path)))[0]
                converted = os.path.join(tmp_dir, basename + '.docx')
                if os.path.exists(converted):
                    print(f"[INFO] LibreOffice 转换成功")
                    return converted
        except Exception as e:
            print(f"[WARN] LibreOffice ({sp}) 转换失败: {e}")

    try:
        import win32com.client
        import pythoncom
        pythoncom.CoInitialize()
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        d = word.Documents.Open(str(doc_path))
        out = os.path.join(tmp_dir, 'converted.docx')
        d.SaveAs2(out, FileFormat=16)
        d.Close()
        word.Quit()
        pythoncom.CoUninitialize()
        print(f"[INFO] Word COM 转换成功")
        return out
    except Exception as e:
        print(f"[WARN] Word COM 转换失败: {e}")

    return None


# ===================================================================
# 2. 提取模板结构概览
# ===================================================================

def extract_template_overview(doc, max_paras=None):
    """提取模板的结构化概览"""
    overview = {"paragraphs": [], "tables": [], "headers": [], "footers": []}

    for i, p in enumerate(doc.paragraphs):
        if max_paras and i >= max_paras:
            break
        text = p.text.strip()
        if not text:
            continue
        try:
            style = p.style.name if p.style else ''
        except Exception:
            style = ''
        if 'toc' in (style or '').lower():
            continue
        overview["paragraphs"].append({
            "index": i, "style": style, "text": text[:500]
        })

    for ti, t in enumerate(doc.tables):
        rows = []
        for row in t.rows:
            cells = []
            for c in row.cells:
                cell_text = c.text.strip().replace('\n', ' | ')
                cells.append(cell_text[:200])
            rows.append(cells)
        overview["tables"].append({"index": ti, "rows": rows})

    for si, sec in enumerate(doc.sections):
        for hf_name in ['header', 'first_page_header', 'even_page_header']:
            hf = getattr(sec, hf_name, None)
            if not hf:
                continue
            texts = []
            for p in hf.paragraphs:
                if p.text.strip():
                    texts.append(p.text.strip())
            for t in hf.tables:
                for row in t.rows:
                    for c in row.cells:
                        if c.text.strip():
                            texts.append(c.text.strip().replace('\n', ' | '))
            if texts:
                overview["headers"].append({
                    "section": si, "type": hf_name, "text": " | ".join(texts)[:500]
                })
        for hf_name in ['footer', 'first_page_footer', 'even_page_footer']:
            hf = getattr(sec, hf_name, None)
            if not hf:
                continue
            texts = []
            for p in hf.paragraphs:
                if p.text.strip():
                    texts.append(p.text.strip())
            if texts:
                overview["footers"].append({
                    "section": si, "type": hf_name, "text": " | ".join(texts)[:300]
                })

    return overview


def format_overview_for_llm(overview):
    lines = []
    lines.append("=== 模板段落（非空，已跳过目录） ===")
    for p in overview["paragraphs"]:
        lines.append(f"[P{p['index']}] ({p['style']}) {p['text']}")
    lines.append("")
    lines.append("=== 模板表格 ===")
    for t in overview["tables"]:
        lines.append(f"--- Table {t['index']} ({len(t['rows'])} rows) ---")
        for ri, row in enumerate(t['rows']):
            lines.append(f"  T{t['index']}.R{ri}: {row}")
    lines.append("")
    lines.append("=== 页眉 ===")
    for h in overview["headers"]:
        lines.append(f"[H{h['section']}.{h['type']}] {h['text']}")
    lines.append("")
    lines.append("=== 页脚 ===")
    for f in overview["footers"]:
        lines.append(f"[F{f['section']}.{f['type']}] {f['text']}")
    return "\n".join(lines)


def format_survey_for_llm(survey):
    field_labels = {
        'sv_company_name': '公司名称', 'sv_cert_other': '其他证书',
        'sv_chairman': '董事长', 'sv_legal_rep': '法人代表',
        'sv_gm': '总经理', 'sv_deputy_gm': '副总经理',
        'sv_mgmt_rep': '管理者代表', 'sv_products': '体系覆盖产品',
        'sv_process_flow': '生产流程', 'sv_location': '地理位置',
        'sv_area': '占地面积', 'sv_building_area': '建筑面积',
        'sv_staff_total': '正式员工人数', 'sv_staff_mgmt': '管理技术人员数',
        'sv_equipment': '设备情况', 'sv_customers': '主要客户',
        'sv_address': '公司地址', 'sv_phone': '电话', 'sv_fax': '传真',
        'sv_mobile': '手机', 'sv_purpose': '公司宗旨/经营理念',
        'sv_quality_policy': '质量方针', 'sv_quality_goal': '质量目标',
        'sv_design_dev': '有无设计开发', 'sv_filler_name': '填写人',
        'sv_certs': '已有证书', 'sv_org': '机构设置',
    }
    lines = ["=== 用户体系调研数据 ==="]
    for key, label in field_labels.items():
        val = survey.get(key, '')
        if isinstance(val, (list, dict)):
            val = json.dumps(val, ensure_ascii=False)
        if val:
            lines.append(f"{label}（{key}）：{val}")
        else:
            lines.append(f"{label}（{key}）：[未填写]")
    return "\n".join(lines)


# ===================================================================
# 3. 构造 LLM 提示词（NDJSON 流式输出）
# ===================================================================

def build_llm_prompt(overview_text, survey_text, template_filename):
    today = datetime.now()
    today_str = f"{today.year}年{today.month}月{today.day}日"
    year_str = str(today.year)

    system = (
        "你是程序文件智能生成助手。你会收到一份程序文件模板的结构概览（带段落索引P#、表格T#.R#、页眉H#、页脚F#）"
        "和用户填写的体系调研数据。你的任务是根据调研数据，决定模板中哪些位置需要修改，逐条输出修改方案。\n\n"
        "【输出格式 - 极其重要】\n"
        "每条修改方案单独输出为一行 JSON 对象（NDJSON 格式），不要包裹在数组里，不要输出任何其他文字。"
        "每生成一条就立即输出一行，不要等所有方案想完再一起输出。输出完所有方案后，最后一行写：===END===\n\n"
        "示例输出（每行一个 JSON，逐行输出）：\n"
        '{"type":"global_replace","old":"AAA企业","new":"诸暨正和金属有限公司","reason":"整体替换公司名"}\n'
        '{"type":"paragraph","index":5,"new_text":"本程序适用于XXX公司所有...","reason":"替换适用范围"}\n'
        '{"type":"table_cell","table":0,"row":2,"col":3,"new_text":"王大明","reason":"填入总经理姓名"}\n'
        '===END===\n\n'
        "修改类型说明：\n"
        "- paragraph: 把段落 P#index 整段替换为 new_text（保留段落格式）\n"
        "- table_cell: 把表格 T#table 第 row 行第 col 列单元格替换为 new_text\n"
        "- global_replace: 在全文（段落+表格+页眉页脚）中把 old 替换为 new\n"
        "- header_replace: 仅在页眉页脚中把 old 替换为 new\n\n"
        "关键规则：\n"
        "1. 公司名必须整体替换：模板里的 AAA企业、AAA 等都要整体替换为用户填写的公司名。\n"
        "2. 文件编号：模板里的文件编号（如 AAA-XX-QP-XX）中的 AAA 替换为合适的公司简称，保留编号结构不变。\n"
        "3. 部门名称：根据模板所属部门，填入调研数据中对应的部门负责人姓名。\n"
        "4. 公司宗旨/质量方针/质量目标：如模板中有相关章节，用调研数据替换。\n"
        "5. 实施日期：模板里的日期替换为" + today_str + "。\n"
        "6. 人名：总经理/管理者代表/贯标办主任等姓名填入对应签字位置。\n"
        "7. 不要修改 IATF16949/ISO9001 标准条款内容。\n"
        "8. 如果调研数据某字段为空，跳过对应修改。\n"
        "9. 修改方案要全面，覆盖所有该改的位置。\n"
        "10. new_text 中的换行用 \\n 转义。\n"
        "11. 想到一条就立即输出一条。\n"
        "12. 保留模板的结构完整性（但必须替换公司名和敏感信息）：\n"
        "    - 不要修改标题的章节编号（如 1.1、5.2.1 等数字编号）\n"
        "    - 不要修改程序文件的步骤编号（如 5.1、5.2、5.3 等）\n"
        "    - 不要删除或重排模板的章节结构\n"
        "    - 不要修改表格的行列结构（只替换单元格内容）\n"
        "    - 保留模板中的所有标准流程步骤的描述原文不动\n"
        "    【但是】以下内容必须替换，不属于'结构完整性'保护范围：\n"
        "    - 模板中所有出现的 AAA、AAA企业 必须替换为用户公司名\n"
        "    - 标题页/封面/页眉中的公司名必须替换\n"
        "    - 文件编号中的 AAA 必须替换（如 AAA-GM-QP-01 → 正和-GM-QP-01）\n"
        "    - 日期必须替换为当前日期\n"
        "    - 人名必须填入对应位置\n"
        "13. 优先使用 global_replace 做简单替换（如公司名、日期、编号中的AAA），只在需要整段重写时用 paragraph。\n"
        "14. 特别是 global_replace 要覆盖所有 AAA 变体：AAA、AAA企业、山东AAA 等，全部替换为用户公司名。\n"
        "15. 【封面表格填写】可以往封面表格（Table 0）的'实施日期/制订/审查/批准'下方的空单元格填写内容：\n"
        "    - 实施日期下方填入当前日期（如 2026年7月11日）\n"
        "    - 制订/审查/批准下方填入对应人名（从调研数据获取）\n"
        "    - 填写内容简短即可，不要换行\n"
        "16. 【页眉页脚保护】不要删除或清空页眉页脚中的任何内容，只做替换：\n"
        "    - 页眉中的公司名、文件编号、文件名称、版次、类别、页次等必须保留\n"
        "    - 只用 header_replace 替换其中的 AAA/公司名/编号，不要用 paragraph 清空页眉\n"
    )

    user = (
        f"当前日期：{today_str}\n"
        f"模板文件：{template_filename}\n\n"
        f"{survey_text}\n\n"
        f"{overview_text}\n\n"
        "请分析以上模板和调研数据，逐条输出 NDJSON 格式的修改方案，每条一行，最后输出 ===END===。"
    )

    return system, user


# ===================================================================
# 4. 解析 LLM 修改方案
# ===================================================================

def parse_ndjson_line(line):
    line = line.strip()
    if not line or line == '===END===':
        return None
    if line.startswith('```'):
        line = line.lstrip('`').replace('json', '', 1).strip()
    if line.endswith('```'):
        line = line[:-3].strip()
    if not line.startswith('{') or not line.endswith('}'):
        return None
    try:
        return json.loads(line)
    except json.JSONDecodeError:
        fixed = re.sub(r',\s*([}\]])', r'\1', line)
        try:
            return json.loads(fixed)
        except json.JSONDecodeError:
            return None


# ===================================================================
# 5. 应用修改方案到 docx
# ===================================================================

def set_paragraph_text(p, new_text):
    if new_text is None:
        new_text = ''
    if not p.runs:
        p.add_run(str(new_text))
        return
    p.runs[0].text = str(new_text)
    for r in p.runs[1:]:
        r.text = ''


def replace_text_in_paragraph(p, old, new):
    if not p.runs or old not in p.text:
        return False
    full = p.text
    new_full = full.replace(old, new)
    if new_full == full:
        return False
    set_paragraph_text(p, new_full)
    return True


def replace_text_in_cell(cell, old, new):
    changed = False
    for p in cell.paragraphs:
        if replace_text_in_paragraph(p, old, new):
            changed = True
    for t in cell.tables:
        for row in t.rows:
            for c in row.cells:
                if replace_text_in_cell(c, old, new):
                    changed = True
    return changed


def apply_global_replace(doc, old, new):
    if not old or old == new:
        return 0
    count = 0
    for p in doc.paragraphs:
        if old in p.text:
            if replace_text_in_paragraph(p, old, new):
                count += 1
    for t in doc.tables:
        for row in t.rows:
            for c in row.cells:
                if old in c.text:
                    if replace_text_in_cell(c, old, new):
                        count += 1
    for sec in doc.sections:
        for hf in [sec.header, sec.first_page_header, sec.even_page_header,
                   sec.footer, sec.first_page_footer, sec.even_page_footer]:
            if not hf:
                continue
            for p in hf.paragraphs:
                if old in p.text:
                    if replace_text_in_paragraph(p, old, new):
                        count += 1
            for t in hf.tables:
                for row in t.rows:
                    for c in row.cells:
                        if old in c.text:
                            if replace_text_in_cell(c, old, new):
                                count += 1
    return count


def apply_header_replace(doc, old, new):
    if not old or old == new:
        return 0
    count = 0
    for sec in doc.sections:
        for hf in [sec.header, sec.first_page_header, sec.even_page_header,
                   sec.footer, sec.first_page_footer, sec.even_page_footer]:
            if not hf:
                continue
            for p in hf.paragraphs:
                if old in p.text:
                    if replace_text_in_paragraph(p, old, new):
                        count += 1
            for t in hf.tables:
                for row in t.rows:
                    for c in row.cells:
                        if old in c.text:
                            if replace_text_in_cell(c, old, new):
                                count += 1
    return count


def apply_paragraph_replace(doc, index, new_text):
    if index < 0 or index >= len(doc.paragraphs):
        print(f"[WARN] 段落索引 {index} 越界")
        return False
    p = doc.paragraphs[index]
    # 【节分界段落保护】如果段落含 sectPr（节分界符），禁止修改内容
    # 节分界段落必须保持空白，填入内容会导致节属性错乱、页眉引用丢失、空白页
    from docx.oxml.ns import qn
    pPr = p._element.find(qn('w:pPr'))
    if pPr is not None and pPr.find(qn('w:sectPr')) is not None:
        print(f"[INFO] 跳过节分界段落 P{index}（含 sectPr，禁止修改）")
        return False
    set_paragraph_text(p, new_text)
    return True


def apply_table_cell_replace(doc, table_idx, row_idx, col_idx, new_text):
    if table_idx < 0 or table_idx >= len(doc.tables):
        print(f"[WARN] 表格索引 {table_idx} 越界")
        return False
    t = doc.tables[table_idx]
    if row_idx < 0 or row_idx >= len(t.rows):
        print(f"[WARN] 行索引 {row_idx} 越界")
        return False
    row = t.rows[row_idx]
    if col_idx < 0 or col_idx >= len(row.cells):
        print(f"[WARN] 列索引 {col_idx} 越界")
        return False
    cell = row.cells[col_idx]
    if cell.paragraphs:
        set_paragraph_text(cell.paragraphs[0], new_text)
        for p in cell.paragraphs[1:]:
            for r in list(p.runs):
                r.text = ''
    else:
        cell.text = str(new_text)
    return True


def apply_modifications(doc, modifications):
    stats = {
        'paragraph': 0, 'table_cell': 0, 'global_replace': 0,
        'header_replace': 0, 'unknown': 0, 'failed': 0,
    }
    for i, mod in enumerate(modifications):
        try:
            mod_type = mod.get('type', '')
            reason = (mod.get('reason', '') or '')[:80]
            if mod_type == 'paragraph':
                idx = int(mod.get('index', -1))
                new_text = mod.get('new_text', '')
                if apply_paragraph_replace(doc, idx, new_text):
                    stats['paragraph'] += 1
                    print(f"  [P{idx}] OK {reason}")
                else:
                    stats['failed'] += 1
            elif mod_type == 'table_cell':
                ti = int(mod.get('table', -1))
                ri = int(mod.get('row', -1))
                ci = int(mod.get('col', -1))
                new_text = mod.get('new_text', '')
                if apply_table_cell_replace(doc, ti, ri, ci, new_text):
                    stats['table_cell'] += 1
                    print(f"  [T{ti}.R{ri}.C{ci}] OK {reason}")
                else:
                    stats['failed'] += 1
            elif mod_type == 'global_replace':
                old = mod.get('old', '')
                new = mod.get('new', '')
                n = apply_global_replace(doc, old, new)
                stats['global_replace'] += n
                print(f"  [G] '{old}' -> '{new}' ({n} 处) OK {reason}")
            elif mod_type == 'header_replace':
                old = mod.get('old', '')
                new = mod.get('new', '')
                n = apply_header_replace(doc, old, new)
                stats['header_replace'] += n
                print(f"  [H] '{old}' -> '{new}' ({n} 处) OK {reason}")
            else:
                stats['unknown'] += 1
        except Exception as e:
            stats['failed'] += 1
            print(f"  [ERROR] 修改 #{i} 失败: {e}")
    return stats


def remove_even_page_headers_footers(doc):
    """删除偶数页页眉页脚引用 + 修复 LibreOffice 转换导致的空白页"""
    from docx.oxml.ns import qn
    removed = 0

    # 1. 删除偶数页引用
    for sec in doc.sections:
        sectPr = sec._sectPr
        for ref in sectPr.findall(qn('w:headerReference')):
            ref_type = ref.get(qn('w:type'))
            if ref_type == 'even':
                sectPr.remove(ref)
                removed += 1
                print(f"[INFO] 删除偶数页页眉引用")
        for ref in sectPr.findall(qn('w:footerReference')):
            ref_type = ref.get(qn('w:type'))
            if ref_type == 'even':
                sectPr.remove(ref)
                removed += 1
                print(f"[INFO] 删除偶数页页脚引用")

    # 2. 修复空白页：LibreOffice 转 .doc→.docx 时，段落0是空段落+nextPage分节符
    # 这会导致封面后多出一个空白页
    # 修复：把段落0的 sectPr 移到段落1，然后删除段落0
    if len(doc.paragraphs) >= 2:
        p0 = doc.paragraphs[0]
        p1 = doc.paragraphs[1]
        pPr0 = p0._element.find(qn('w:pPr'))
        pPr1 = p1._element.find(qn('w:pPr'))

        if pPr0 is not None:
            sectPr0 = pPr0.find(qn('w:sectPr'))
            if sectPr0 is not None and not p0.text.strip():
                # 段落0有分节符且内容为空
                if pPr1 is None:
                    pPr1 = p1._element.makeelement(qn('w:pPr'), {})
                    p1._element.insert(0, pPr1)
                existing = pPr1.find(qn('w:sectPr'))
                if existing is None:
                    pPr1.append(sectPr0)
                    print(f"[INFO] 将段落0的分节符移动到段落1（修复空白页）")
                p0._element.getparent().remove(p0._element)
                print(f"[INFO] 删除空段落0（修复空白页）")

    return removed
