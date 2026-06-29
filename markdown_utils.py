"""
Markdown 解析与处理工具方法。

包含：
- PENDING_POINT_COLUMNS: 待确认点表格列名
- is_markdown_separator_row: 判断 Markdown 表格分隔行
- parse_pending_points_from_markdown: 从 Markdown 解析待确认点清单
- remove_pending_points_section: 从 Markdown 中移除待确认点清单部分
- pending_points_to_llm_text: 将待确认点表格转成传给大模型的文本
- pending_points_to_markdown: 将待确认点转成 Markdown 表格
- get_pending_answer_key: 生成稳定的 widget key
- extract_sql_section_from_test_result: 从测试结果中提取 SQL 章节
- extract_sql_code_blocks: 提取 Markdown 中的 sql 代码块
- strip_markdown_fence: 去除 Markdown 代码块标记
- escape_sql_block_comment: 避免 */ 破坏 SQL 多行注释
"""

import re
import hashlib

import streamlit as st


PENDING_POINT_COLUMNS = [
    "待确认编号",
    "待确认问题",
    "影响范围",
    "建议用户补充内容",
    "用户补充说明"
]


def is_markdown_separator_row(cells: list[str]) -> bool:
    """
    判断 Markdown 表格中的分隔行，例如：
    |---|---|---|
    """
    if not cells:
        return False

    for cell in cells:
        cell = cell.strip()
        if not re.fullmatch(r":?-{3,}:?", cell):
            return False

    return True


def parse_pending_points_from_markdown(markdown_text: str) -> list[dict]:
    """
    从第一步 AI 输出的 Markdown 中解析"待确认点清单"表格。
    兼容：
    - ## 三、待确认点清单
    - ### 三、待确认点清单
    - 三、待确认点清单
    - ## 三、待确认问题清单
    - ## 三、待确认事项清单
    """
    if not markdown_text:
        return []

    lines = markdown_text.splitlines()

    # 1. 找到待确认点章节起始行，不强依赖"## 三、待确认点清单"
    start_idx = None

    for i, line in enumerate(lines):
        clean_line = line.strip()

        if "待确认" in clean_line and (
            "清单" in clean_line or "问题" in clean_line or "事项" in clean_line
        ):
            start_idx = i
            break

    if start_idx is None:
        return []

    # 2. 找到下一个章节标题作为结束位置
    end_idx = len(lines)

    for j in range(start_idx + 1, len(lines)):
        clean_line = lines[j].strip()

        # 匹配类似：## 四、xxx / ### 四、xxx / 四、xxx
        if re.match(r"^\s*#{0,6}\s*[一二三四五六七八九十]+[、.．]\s*", clean_line):
            if "待确认" not in clean_line:
                end_idx = j
                break

    section_text = "\n".join(lines[start_idx:end_idx])

    # 3. 如果明确是无，则返回空
    if re.search(r"(?m)^\s*无\s*$", section_text.strip()):
        return []

    # 4. 提取 Markdown 表格行
    table_lines = []

    for line in section_text.splitlines():
        line = line.strip()
        if line.startswith("|") and "|" in line:
            table_lines.append(line)

    if not table_lines:
        return []

    header = None
    rows = []

    for line in table_lines:
        cells = [cell.strip() for cell in line.strip().strip("|").split("|")]

        if is_markdown_separator_row(cells):
            continue

        if header is None:
            header = cells
            continue

        if not cells:
            continue

        if len(cells) < len(header):
            cells = cells + [""] * (len(header) - len(cells))

        row_dict = {}

        for col in PENDING_POINT_COLUMNS:
            if col in header:
                idx = header.index(col)
                row_dict[col] = cells[idx] if idx < len(cells) else ""
            else:
                row_dict[col] = ""

        # 兼容列名变化
        if not row_dict["待确认编号"]:
            for alt_col in ["编号", "问题编号", "待确认点编号"]:
                if alt_col in header:
                    idx = header.index(alt_col)
                    row_dict["待确认编号"] = cells[idx] if idx < len(cells) else ""
                    break

        if not row_dict["待确认问题"]:
            for alt_col in ["待确认点", "问题", "确认问题", "待确认事项"]:
                if alt_col in header:
                    idx = header.index(alt_col)
                    row_dict["待确认问题"] = cells[idx] if idx < len(cells) else ""
                    break

        if not row_dict["影响范围"]:
            for alt_col in ["影响", "影响说明"]:
                if alt_col in header:
                    idx = header.index(alt_col)
                    row_dict["影响范围"] = cells[idx] if idx < len(cells) else ""
                    break

        if not row_dict["建议用户补充内容"]:
            for alt_col in ["建议补充内容", "需补充内容", "用户需补充内容"]:
                if alt_col in header:
                    idx = header.index(alt_col)
                    row_dict["建议用户补充内容"] = cells[idx] if idx < len(cells) else ""
                    break

        # 用户补充说明保持空，等待用户填写
        row_dict["用户补充说明"] = ""

        # 过滤空行
        if any(str(row_dict[col]).strip() for col in PENDING_POINT_COLUMNS[:-1]):
            rows.append(row_dict)

    return rows


def remove_pending_points_section(markdown_text: str) -> str:
    """
    从 AI 输出中移除"待确认点清单"部分。
    页面上改用可编辑表格展示待确认点，避免重复显示不可编辑 Markdown 表格。
    兼容：
    - ## 三、待确认点清单
    - ### 三、待确认问题清单
    - 三、待确认事项
    """
    if not markdown_text:
        return ""

    lines = markdown_text.splitlines()

    start_idx = None

    for i, line in enumerate(lines):
        clean_line = line.strip()
        if "待确认" in clean_line and (
            "清单" in clean_line or "问题" in clean_line or "事项" in clean_line
        ):
            start_idx = i
            break

    if start_idx is None:
        return markdown_text.strip()

    end_idx = len(lines)

    for j in range(start_idx + 1, len(lines)):
        clean_line = lines[j].strip()

        if re.match(r"^\s*#{0,6}\s*[一二三四五六七八九十]+[、.．]\s*", clean_line):
            if "待确认" not in clean_line:
                end_idx = j
                break

    new_lines = lines[:start_idx] + lines[end_idx:]

    return "\n".join(new_lines).strip()


def pending_points_to_llm_text(rows: list[dict]) -> str:
    """
    将可编辑待确认点表格转换成传给第二步大模型的文本。
    """
    if not rows:
        return "未解析到待确认点清单，或第一步未产生待确认点。"

    lines = []

    for row in rows:
        pending_id = str(row.get("待确认编号", "")).strip()
        question = str(row.get("待确认问题", "")).strip()
        impact = str(row.get("影响范围", "")).strip()
        suggestion = str(row.get("建议用户补充内容", "")).strip()
        answer = str(row.get("用户补充说明", "")).strip()

        if not pending_id and not question:
            continue

        lines.append(
            f"""
待确认编号：{pending_id if pending_id else "未提供"}
待确认问题：{question if question else "未提供"}
影响范围：{impact if impact else "未提供"}
建议用户补充内容：{suggestion if suggestion else "未提供"}
用户补充说明：{answer if answer else "未补充"}
"""
        )

    return "\n".join(lines).strip() if lines else "无待确认点。"


def pending_points_to_markdown(rows: list[dict]) -> str:
    """
    将可编辑后的待确认点清单转换成 Markdown，主要用于下载文件。
    """
    if not rows:
        return "## 三、待确认点清单\n\n无"

    lines = []
    lines.append("## 三、待确认点清单")
    lines.append("")
    lines.append("| 待确认编号 | 待确认问题 | 影响范围 | 建议用户补充内容 | 用户补充说明 |")
    lines.append("|---|---|---|---|---|")

    for row in rows:
        line = "| {id} | {question} | {impact} | {suggestion} | {answer} |".format(
            id=str(row.get("待确认编号", "")).replace("|", "\\|"),
            question=str(row.get("待确认问题", "")).replace("|", "\\|"),
            impact=str(row.get("影响范围", "")).replace("|", "\\|"),
            suggestion=str(row.get("建议用户补充内容", "")).replace("|", "\\|"),
            answer=str(row.get("用户补充说明", "")).replace("|", "\\|"),
        )
        lines.append(line)

    return "\n".join(lines)


def get_pending_answer_key(row: dict, index: int) -> str:
    """
    为每个待确认点的"用户补充说明"生成稳定的 widget key。
    key 中包含 pending_points_editor_version，重新生成第一步后会自动换一批 key，避免旧输入串到新结果。
    """
    version = st.session_state.get("pending_points_editor_version", 0)

    raw = (
        f"{version}_"
        f"{index}_"
        f"{row.get('待确认编号', '')}_"
        f"{row.get('待确认问题', '')}"
    )

    md5 = hashlib.md5()
    md5.update(raw.encode("utf-8"))

    return f"pending_answer_{md5.hexdigest()}"


def extract_sql_section_from_test_result(result_text: str) -> str:
    """
    从测试用例结果中提取"三、SQL 校验脚本"章节。
    """
    if not result_text:
        return ""

    pattern = r"(?m)^#{1,6}\s*三[、.．]\s*SQL\s*校验脚本\s*$"
    match = re.search(pattern, result_text)

    if not match:
        return result_text.strip()

    start = match.end()

    # 找下一个同级或任意后续章节标题，例如 ## 四、说明
    next_match = re.search(
        r"(?m)^#{1,6}\s*[四五六七八九十][、.．]\s*",
        result_text[start:]
    )

    if next_match:
        end = start + next_match.start()
        return result_text[start:end].strip()

    return result_text[start:].strip()


def extract_sql_code_blocks(markdown_text: str) -> list[str]:
    """
    提取 Markdown 中的 sql 代码块。
    """
    if not markdown_text:
        return []

    blocks = re.findall(
        r"```(?:sql|SQL)?\s*\n(.*?)```",
        markdown_text,
        flags=re.DOTALL
    )

    return [block.strip() for block in blocks if block.strip()]


def strip_markdown_fence(text: str) -> str:
    """
    去除 Markdown 代码块标记。
    """
    if not text:
        return ""

    text = re.sub(r"```(?:sql|SQL)?", "", text)
    text = text.replace("```", "")

    return text.strip()


def escape_sql_block_comment(text: str) -> str:
    """
    避免测试用例说明中的 */ 破坏 SQL 多行注释。
    """
    if not text:
        return ""

    return text.replace("*/", "* /")
