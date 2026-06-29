"""
UI 组件与页面辅助方法。

包含：
- inject_custom_css: 注入页面美化 CSS
- render_markdown_in_scroll_box: 滚动容器展示长 Markdown
- render_step_progress: 顶部流程进度展示 + 步骤点击切换
- render_page_header: 每个步骤顶部的说明卡片
- build_vscode_runnable_sql_download_content: 构造可执行的 .sql 文件内容
- _split_sections_by_heading: 按 heading 关键词拆分 result_text
- render_test_case_result_with_download: 渲染测试用例结果
- render_pending_points_data_editor: 展示待确认点可编辑表格
- sync_pending_points_from_widgets: 从 widget 同步用户补充说明
- render_table_schema_uploader: 表结构组件（ODPS 拉取 / xlsx 上传）
- go_to_step: 跳转到指定步骤
- get_materials_from_state: 从 session_state 获取所有材料
"""

import re

import pandas as pd
import streamlit as st

from constants import STEP_INPUT, STEP_OPTIONS
from file_utils import get_uploaded_file_id, read_table_schema_xlsx
from markdown_utils import (
    PENDING_POINT_COLUMNS,
    escape_sql_block_comment,
    extract_sql_code_blocks,
    extract_sql_section_from_test_result,
    get_pending_answer_key,
    pending_points_to_llm_text,
    pending_points_to_markdown,
    strip_markdown_fence,
)
from odps_utils import (
    get_odps_entry,
    get_table_schema_text,
    is_partitioned_table,
    preview_table_data,
)


def go_to_step(step_name: str) -> None:
    st.session_state["current_step"] = step_name
    st.rerun()


def get_materials_from_state() -> dict:
    return {
        "prd_text": st.session_state.get("prd_text", ""),
        "meeting_notes": st.session_state.get("meeting_notes", ""),
        "result_table_schema": st.session_state.get("result_table_schema", ""),
        "source_table_schema": st.session_state.get("source_table_schema", ""),
        "dev_code": st.session_state.get("dev_code", ""),
    }


def inject_custom_css() -> None:
    """
    注入页面美化 CSS — Notion留白 + AntPro轻量化 + Vercel工程质感。
    """
    st.markdown(
        """
<style>
/* ===== 全局 ===== */
.block-container {
    padding-top: 2.5rem;
    padding-bottom: 5rem;
    max-width: 1200px;
}

/* 字体层级 */
h1 { font-size: 20px !important; font-weight: 600 !important; letter-spacing: -0.02em; color: #0F172A; }
h2 { font-size: 16px !important; font-weight: 600 !important; letter-spacing: -0.01em; color: #0F172A; }
h3 { font-size: 14px !important; font-weight: 600 !important; color: #334155; }

/* 正文 */
p, li {
    font-size: 14px !important;
    line-height: 1.7;
    color: #334155;
}

/* ===== 顶部标题区 — Notion式白底，下划线分隔 ===== */
.app-hero {
    padding: 0 0 12px 0;
    border-radius: 0;
    background: transparent;
    border-bottom: 2px solid #0F172A;
    margin-bottom: 28px;
}

.app-hero-title {
    font-size: 24px;
    font-weight: 700;
    color: #0F172A;
    margin-bottom: 4px;
    letter-spacing: -0.03em;
}

.app-hero-desc {
    font-size: 13px;
    color: #64748B;
    line-height: 1.5;
}

/* ===== 步骤进度卡片 — 克制，蓝色仅点缀 ===== */
.step-card {
    padding: 14px 12px;
    border-radius: 6px;
    border: 1px solid #E2E8F0;
    background: #FFFFFF;
    text-align: center;
    min-height: 64px;
    transition: all 0.15s ease;
}

.step-card-active {
    border: 1px solid #2563EB;
    background: #EFF6FF;
}

.step-card-done {
    border: 1px solid #E2E8F0;
    background: #FAFAFA;
}

.step-card-title {
    font-size: 13px;
    font-weight: 500;
    color: #64748B;
    line-height: 1.4;
}

.step-card-active .step-card-title {
    color: #2563EB;
    font-weight: 600;
}

.step-card-done .step-card-title {
    color: #94A3B8;
}

.step-card-status {
    margin-top: 4px;
    font-size: 11px;
    color: #CBD5E1;
    font-weight: 400;
}

.step-card-active .step-card-status {
    color: #60A5FA;
}

.step-card-done .step-card-status {
    color: #CBD5E1;
}

/* ===== 步骤说明卡片 — 去边框，Notion式底线 ===== */
.page-section-card {
    padding: 0 0 16px 0;
    border-radius: 0;
    background: transparent;
    border: none;
    border-bottom: 1px solid #F1F5F9;
    margin-bottom: 24px;
}

.page-section-title {
    font-size: 18px;
    font-weight: 600;
    color: #0F172A;
    margin-bottom: 2px;
    letter-spacing: -0.01em;
}

.page-section-desc {
    font-size: 13px;
    color: #94A3B8;
    line-height: 1.5;
}

/* ===== 按钮 — 默认中性灰，蓝色仅primary ===== */
.stButton > button,
.stDownloadButton > button {
    border-radius: 6px !important;
    min-height: 36px;
    font-weight: 500 !important;
    font-size: 13px !important;
    transition: all 0.12s ease;
    border: 1px solid #E2E8F0 !important;
    background: #FFFFFF !important;
    color: #334155 !important;
}

.stButton > button:hover,
.stDownloadButton > button:hover {
    border-color: #CBD5E1 !important;
    background: #F8FAFC !important;
    color: #0F172A !important;
}

/* primary 按钮 — 蓝底白字，强制覆盖所有选择器 */
.stButton > button[kind="primary"],
div[data-testid="stButton"] > button[kind="primary"],
.stButton > button[data-testid="stBaseButton-primary"],
div[data-testid="stButton"] button[data-testid="stBaseButton-primary"],
button[kind="primary"],
button[data-testid="stBaseButton-primary"],
.stButton button[kind="primary"],
section[data-testid="stMain"] button[kind="primary"],
div[class*="stButton"] button[kind="primary"] {
    background: #2563EB !important;
    color: #FFFFFF !important;
    border: 1px solid #2563EB !important;
    font-weight: 500 !important;
}

.stButton > button[kind="primary"]:hover,
div[data-testid="stButton"] > button[kind="primary"]:hover,
.stButton > button[data-testid="stBaseButton-primary"]:hover,
button[kind="primary"]:hover,
button[data-testid="stBaseButton-primary"]:hover,
.stButton button[kind="primary"]:hover,
section[data-testid="stMain"] button[kind="primary"]:hover,
div[class*="stButton"] button[kind="primary"]:hover {
    background: #1D4ED8 !important;
    border-color: #1D4ED8 !important;
    color: #FFFFFF !important;
}

/* 兜底：任何带 primary 的按钮强制白字 */
button[class*="primary"],
button[class*="Primary"],
.st-emotion-cache-1dumvfu,
button.st-emotion-cache-1dumvfu,
.stButton button.st-emotion-cache-1dumvfu {
    color: #FFFFFF !important;
    background: #2563EB !important;
    border-color: #2563EB !important;
}

/* ===== 输入框 ===== */
textarea, input {
    border-radius: 6px !important;
    font-size: 14px !important;
}

/* ===== expander — 淡底色无边框 ===== */
div[data-testid="stExpander"] {
    border-radius: 6px !important;
    border: 1px solid #F1F5F9 !important;
    background: #FAFAFA !important;
}

/* ===== sidebar ===== */
section[data-testid="stSidebar"] {
    background: #FAFAFA;
    border-right: 1px solid #F1F5F9;
}

section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3 {
    color: #0F172A;
    font-size: 12px !important;
    font-weight: 600 !important;
    text-transform: uppercase;
    letter-spacing: 0.06em;
}

/* ===== data editor ===== */
div[data-testid="stDataFrame"] {
    border-radius: 6px;
}

/* ===== 分割线 — 极淡 ===== */
hr {
    border: none;
    border-top: 1px solid #F1F5F9;
    margin-top: 2rem;
    margin-bottom: 2rem;
}

/* ===== alert — 去边框 ===== */
div[data-testid="stAlert"] {
    border-radius: 6px;
    border: none !important;
}

/* ===== 代码块 ===== */
pre {
    border-radius: 6px !important;
    font-size: 13px !important;
}

/* ===== markdown 表格 ===== */
table {
    border-radius: 6px;
    overflow: hidden;
    border: 1px solid #E2E8F0 !important;
}

th {
    font-weight: 600 !important;
    font-size: 12px !important;
    background: #F8FAFC !important;
    color: #475569 !important;
}

td {
    font-size: 13px !important;
    color: #334155 !important;
}

/* ===== tab ===== */
div[data-testid="stTabs"] {
    border-radius: 6px;
}

/* ===== container(可滚动) — 细边框 ===== */
div[data-testid="stVerticalBlockBorderWrapper"] {
    border: 1px solid #F1F5F9 !important;
    border-radius: 6px !important;
}

/* element 间距 */
.element-container {
    margin-bottom: 8px;
}

/* ===== 强制蓝底 primary 按钮内部文字为白色 ===== */
.stButton button[kind="primary"] *,
div[data-testid="stButton"] button[kind="primary"] *,
button[kind="primary"] *,
.stButton button[data-testid="stBaseButton-primary"] *,
div[data-testid="stButton"] button[data-testid="stBaseButton-primary"] *,
button[data-testid="stBaseButton-primary"] *,
section[data-testid="stMain"] button[kind="primary"] *,
div[class*="stButton"] button[kind="primary"] * {
    color: #FFFFFF !important;
    fill: #FFFFFF !important;
}
/* primary 按钮 hover 时，内部文字也保持白色 */
.stButton button[kind="primary"]:hover *,
div[data-testid="stButton"] button[kind="primary"]:hover *,
button[kind="primary"]:hover *,
.stButton button[data-testid="stBaseButton-primary"]:hover *,
div[data-testid="stButton"] button[data-testid="stBaseButton-primary"]:hover *,
button[data-testid="stBaseButton-primary"]:hover *,
section[data-testid="stMain"] button[kind="primary"]:hover *,
div[class*="stButton"] button[kind="primary"]:hover * {
    color: #FFFFFF !important;
    fill: #FFFFFF !important;
}

/* element 间距 */
.element-container {
    margin-bottom: 8px;
}
/* ===== 顶部步骤导航按钮 ===== */
div[data-testid="stHorizontalBlock"] div[data-testid="stButton"] > button {
    min-height: 44px;
    border-radius: 8px !important;
    font-size: 13px !important;
    font-weight: 600 !important;
}
/* 顶部当前步骤 primary 按钮文字强制白色 */
div[data-testid="stHorizontalBlock"] div[data-testid="stButton"] > button[kind="primary"],
div[data-testid="stHorizontalBlock"] div[data-testid="stButton"] > button[data-testid="stBaseButton-primary"] {
    background: #2563EB !important;
    border-color: #2563EB !important;
    color: #FFFFFF !important;
}
div[data-testid="stHorizontalBlock"] div[data-testid="stButton"] > button[kind="primary"] *,
div[data-testid="stHorizontalBlock"] div[data-testid="stButton"] > button[data-testid="stBaseButton-primary"] * {
    color: #FFFFFF !important;
}

/* ===== 允许文字选中复制，禁止 Streamlit 默认右键菜单拦截 ===== */
/* 恢复用户选中和右键能力 */
[data-testid="stAppViewContainer"],
[data-testid="stAppViewContainer"] * {
    -webkit-user-select: text !important;
    user-select: text !important;
}

/* 禁止 Streamlit 拦截 contextmenu 事件 */
div[data-testid="stAppViewBlockContainer"] {
    pointer-events: auto !important;
}

</style>
        """,
        unsafe_allow_html=True
    )


def render_markdown_in_scroll_box(
    title: str,
    markdown_text: str,
    height: int = 520,
    expanded: bool = True
) -> None:
    """
    用 expander + 固定高度容器展示长 Markdown，避免页面过长。
    """
    if not markdown_text:
        return

    with st.expander(title, expanded=expanded):
        with st.container(height=height, border=True):
            st.markdown(markdown_text)


def render_step_progress() -> None:
    """
    顶部流程进度展示 + 步骤点击切换。
    """
    current_step = st.session_state.get("current_step", STEP_INPUT)

    if current_step in STEP_OPTIONS:
        current_index = STEP_OPTIONS.index(current_step)
    else:
        current_index = 0

    progress_value = (current_index + 1) / len(STEP_OPTIONS)
    st.progress(progress_value)

    cols = st.columns(len(STEP_OPTIONS))

    for index, step in enumerate(STEP_OPTIONS):
        if index < current_index:
            status = "已完成"
            status_icon = "—"
            button_type = "secondary"
        elif index == current_index:
            status = "当前步骤"
            status_icon = "●"
            button_type = "primary"
        else:
            status = "待处理"
            status_icon = "○"
            button_type = "secondary"

        with cols[index]:
            clicked = st.button(
                step,
                key=f"top_step_nav_{index}",
                type=button_type,
                use_container_width=True,
                help=f"{status_icon} {status}"
            )

            st.markdown(
                f"""
<div style="
    text-align:center;
    font-size:11px;
    color:#94A3B8;
    margin-top:-4px;
    margin-bottom:8px;
">
    {status_icon} {status}
</div>
                """,
                unsafe_allow_html=True
            )

            if clicked:
                st.session_state["current_step"] = step
                st.rerun()

    st.write("")


def render_page_header(title: str, desc: str, icon: str = "") -> None:
    """
    每个步骤顶部的说明卡片。
    """
    title_html = f"{icon} {title}" if icon else title
    st.markdown(
        f"""
<div class="page-section-card">
    <div class="page-section-title">{title_html}</div>
    <div class="page-section-desc">{desc}</div>
</div>
        """,
        unsafe_allow_html=True
    )


def build_vscode_runnable_sql_download_content(result_text: str) -> str:
    """
    构造可在 VSCode 中打开并执行的 .sql 文件内容。

    文件结构：
    1. 前半部分：测试用例和关注点，以 SQL 多行注释形式保留。
    2. 后半部分：真正可执行 SQL。
    """
    if not result_text:
        return ""

    # 找到 SQL 章节位置
    sql_heading_pattern = r"(?m)^#{1,6}\s*三[、.．]\s*SQL\s*校验脚本\s*$"
    match = re.search(sql_heading_pattern, result_text)

    if match:
        before_sql_section = result_text[:match.start()].strip()
        sql_section = extract_sql_section_from_test_result(result_text)
    else:
        before_sql_section = ""
        sql_section = result_text.strip()

    # 优先提取 ```sql ... ``` 代码块
    sql_blocks = extract_sql_code_blocks(sql_section)

    if sql_blocks:
        sql_body = "\n\n\n".join(sql_blocks).strip()
    else:
        # 兜底：如果模型没有输出代码块，则去掉 Markdown fence 后直接作为 SQL 内容
        sql_body = strip_markdown_fence(sql_section)

    before_sql_section = escape_sql_block_comment(before_sql_section)

    sql_file_content = f"""/*
=====================================================
数据测试用例说明
=====================================================

以下内容来自 AI 生成的测试关注点和测试用例。
为了保证 .sql 文件可以在 VSCode / SQL 客户端中直接打开运行，
测试用例说明已放入 SQL 注释中。

{before_sql_section if before_sql_section else "无额外测试用例说明。"}

=====================================================
*/

-- =====================================================
-- 数据测试 SQL 校验脚本
-- 使用说明：
-- 1. 请确认 VSCode 已安装对应数据库插件，例如 SQLTools、Hive、Spark SQL、Presto、Trino、MySQL 等。
-- 2. 请确认当前连接的数据源与 SQL 方言一致。
-- 3. 如果 SQL 中存在 TODO 注释，请先替换为真实表名、字段名或分区值。
-- 4. 每段 SQL 可单独选中执行，也可以整体执行。
-- =====================================================

{sql_body.rstrip()}

"""
    return sql_file_content


def _split_sections_by_heading(result_text: str, headings: list) -> dict:
    """
    按 heading 关键词将 result_text 拆分为多段。
    headings 示例: ["一", "二", "三"]
    返回: {"一": "...", "二": "...", "三": "..."}
    每个 value 包含从该标题到下一个标题前的完整内容（含标题行）。
    """
    # 构建正则：匹配 ## 一、 或 ## 一. 或 ## 一． 等
    patterns = {}
    for h in headings:
        patterns[h] = re.compile(
            rf"(?m)^(#{{1,6}})\s*{h}[、.．]"
        )

    # 找到每个标题的位置
    positions = []
    for h, p in patterns.items():
        m = p.search(result_text)
        if m:
            positions.append((m.start(), h, m))
    positions.sort(key=lambda x: x[0])

    sections = {}
    for i, (start, h, m) in enumerate(positions):
        end = positions[i + 1][0] if i + 1 < len(positions) else len(result_text)
        sections[h] = result_text[start:end].strip()

    return sections


def render_test_case_result_with_download(result_text: str) -> None:
    """
    渲染测试用例结果：
    - SQL 下载按钮放在最顶部
    - 一、二、三三个章节各自包在可滚动容器里（st.container border + height）
    - 章节之间用 st.divider() 分隔
    """
    if not result_text:
        return

    sql_download_content = build_vscode_runnable_sql_download_content(result_text)

    sections = _split_sections_by_heading(result_text, ["一", "二", "三"])

    # 下载按钮始终放最顶部
    st.download_button(
        label="下载 SQL 脚本",
        data=sql_download_content,
        file_name="data_test_validation.sql",
        mime="text/plain",
        use_container_width=True
    )

    # 如果没匹配到任何章节标题，退回整段渲染（放在可滚动容器里）
    if not sections:
        with st.container(height=400, border=True):
            st.markdown(result_text)
        return

    # 一、测试关注点 — 可滚动容器
    if "一" in sections:
        with st.container(height=350, border=True):
            st.markdown(sections["一"])

    st.divider()

    # 二、测试用例清单 — 可滚动容器
    if "二" in sections:
        with st.container(height=400, border=True):
            st.markdown(sections["二"])

    st.divider()

    # 三、SQL 校验脚本 — 可滚动容器
    if "三" in sections:
        with st.container(height=450, border=True):
            st.markdown(sections["三"])


def render_pending_points_data_editor() -> None:
    """
    使用 st.data_editor 紧凑展示待确认点，只允许编辑"用户补充说明"列。
    """
    rows = st.session_state.get("pending_points_rows", [])

    if not rows:
        st.success("当前没有阻塞测试设计或 SQL 校验的待确认点，可以继续生成最终版需求提炼表。")
        st.session_state["prd_pending_answers"] = "无待确认点。"
        return

    df = pd.DataFrame(rows)

    for col in PENDING_POINT_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    df = df[PENDING_POINT_COLUMNS].fillna("")

    editor_key = f"pending_points_data_editor_{st.session_state.get('pending_points_editor_version', 0)}"

    edited_df = st.data_editor(
        df,
        key=editor_key,
        hide_index=True,
        use_container_width=True,
        height=min(460, 110 + len(df) * 58),
        disabled=[
            "待确认编号",
            "待确认问题",
            "影响范围",
            "建议用户补充内容"
        ],
        column_config={
            "待确认编号": st.column_config.TextColumn(
                "待确认编号",
                width="small"
            ),
            "待确认问题": st.column_config.TextColumn(
                "待确认问题",
                width="large"
            ),
            "影响范围": st.column_config.TextColumn(
                "影响范围",
                width="medium"
            ),
            "建议用户补充内容": st.column_config.TextColumn(
                "建议用户补充内容",
                width="large"
            ),
            "用户补充说明": st.column_config.TextColumn(
                "用户补充说明，可编辑",
                width="large",
                help="请在这里填写确认结果；如果认为可忽略，也可以不填并点击忽略。"
            ),
        },
        num_rows="fixed"
    )

    st.session_state["pending_points_rows"] = edited_df.fillna("").to_dict("records")

    st.session_state["prd_pending_answers"] = pending_points_to_llm_text(
        st.session_state["pending_points_rows"]
    )


def sync_pending_points_from_widgets() -> None:
    """
    从每个待确认点对应的 text_area widget 中同步用户补充说明，
    并更新 prd_pending_answers，供第二步和测试用例生成使用。
    """
    rows = st.session_state.get("pending_points_rows", [])

    if not rows:
        st.session_state["prd_pending_answers"] = "无待确认点。"
        return

    for index, row in enumerate(rows):
        key = get_pending_answer_key(row, index)

        if key in st.session_state:
            row["用户补充说明"] = str(st.session_state[key]).strip()
        else:
            row["用户补充说明"] = str(row.get("用户补充说明", "")).strip()

    st.session_state["pending_points_rows"] = rows
    st.session_state["prd_pending_answers"] = pending_points_to_llm_text(rows)


def render_table_schema_uploader(title: str, state_prefix: str) -> str:
    """
    渲染表结构组件：支持下拉选择"ODPS 拉取"或"xlsx 上传"。
    ODPS 模式：逐行添加表名，一键批量拉取，失败跳过可重试。
    每行一张表，行尾跟分区填写框。

    参数：
    title: 页面显示名称，例如：结果表表结构、源表表结构
    state_prefix: session_state 前缀，例如：result_schema、source_schema

    返回：
    拼接后的表结构文本，用于传给大模型。
    """
    items_key = f"{state_prefix}_items"
    uploader_version_key = f"{state_prefix}_uploader_version"
    mode_key = f"{state_prefix}_mode"

    if items_key not in st.session_state:
        st.session_state[items_key] = []
    if uploader_version_key not in st.session_state:
        st.session_state[uploader_version_key] = 0
    if mode_key not in st.session_state:
        st.session_state[mode_key] = "odps" if get_odps_entry(
            st.session_state.get("odps_ak", "").strip(),
            st.session_state.get("odps_sk", "").strip(),
            st.session_state.get("odps_project", "").strip(),
            st.session_state.get("odps_endpoint", "").strip(),
        ) else "xlsx"

    st.subheader(title)

    # 模式下拉
    _odps_available = get_odps_entry(
        st.session_state.get("odps_ak", "").strip(),
        st.session_state.get("odps_sk", "").strip(),
        st.session_state.get("odps_project", "").strip(),
        st.session_state.get("odps_endpoint", "").strip(),
    ) is not None

    if _odps_available:
        _mode_options = ["odps", "xlsx"]
        _mode_labels = {"odps": "ODPS 拉取", "xlsx": "xlsx 上传"}
        _mode = st.radio(
            f"选择{title}输入方式",
            options=_mode_options,
            format_func=lambda x: _mode_labels[x],
            key=mode_key,
            horizontal=True
        )
    else:
        _mode = "xlsx"
        st.caption("未配置 ODPS 连接，仅支持 xlsx 上传。")

    # ===== xlsx 上传模式 =====
    if _mode == "xlsx":
        uploaded_files = st.file_uploader(
            f"上传 {title} xlsx 文件，可多选",
            type=["xlsx"],
            accept_multiple_files=True,
            key=f"{state_prefix}_uploader_{st.session_state[uploader_version_key]}"
        )

        if uploaded_files:
            existing_ids = {item["id"] for item in st.session_state[items_key]}
            for uploaded_file in uploaded_files:
                file_bytes = uploaded_file.getvalue()
                file_id = get_uploaded_file_id(uploaded_file.name, file_bytes)
                if file_id not in existing_ids:
                    schema_text = read_table_schema_xlsx(file_bytes, uploaded_file.name)
                    st.session_state[items_key].append({
                        "id": file_id,
                        "name": uploaded_file.name,
                        "source": "xlsx",
                        "schema_text": schema_text,
                        "partition": "",
                        "fetch_status": "ok",
                        "fetch_error": ""
                    })
                    existing_ids.add(file_id)

    # ===== ODPS 拉取模式 =====
    elif _mode == "odps":
        # 批量添加表名（多行文本框，每行一个表名）
        _c_add_input, _c_add_btn = st.columns([4, 1])
        with _c_add_input:
            _add_tbl_text = st.text_area(
                "输入表名，每行一个（支持 项目名.表名）",
                key=f"{state_prefix}_add_input_{st.session_state[uploader_version_key]}",
                height=80,
                placeholder="例如：\nods_project.ods_order_detail_di\ndwd_project.dwd_order_d\nads_project.ads_order_summary_df",
                label_visibility="collapsed"
            )
        with _c_add_btn:
            if st.button("一键添加", key=f"{state_prefix}_add_btn", use_container_width=True):
                if _add_tbl_text.strip():
                    _existing_names = {x["name"] for x in st.session_state[items_key]}
                    _added_count = 0
                    for _line in _add_tbl_text.strip().splitlines():
                        _tbl_name = _line.strip()
                        if _tbl_name and _tbl_name not in _existing_names:
                            _new_id = f"odps_{_tbl_name}_{len(st.session_state[items_key])}"
                            st.session_state[items_key].append({
                                "id": _new_id,
                                "name": _tbl_name,
                                "source": "odps",
                                "schema_text": "",
                                "partition": "",
                                "fetch_status": "pending",
                                "fetch_error": ""
                            })
                            _existing_names.add(_tbl_name)
                            _added_count += 1
                    if _added_count:
                        st.success(f"已添加 {_added_count} 张表。")
                        # 清空输入框：用回调方式，在 widget 创建前设
                        # 不能在 widget 创建后改 session_state，改用 version 换 key
                        st.session_state[uploader_version_key] += 1
                    st.rerun()

        # 批量填写分区
        if st.session_state[items_key]:
            _c_pt_input, _c_pt_btn = st.columns([4, 1])
            with _c_pt_input:
                _batch_pt_val = st.text_input(
                    "批量填写分区，应用到所有表",
                    key=f"{state_prefix}_batch_pt_input",
                    placeholder="例如：pt='20250101'  留空=无分区",
                    label_visibility="collapsed"
                )
            with _c_pt_btn:
                if st.button("一键填分区", key=f"{state_prefix}_batch_pt_btn", use_container_width=True):
                    _pt_to_set = _batch_pt_val.strip()
                    for _item in st.session_state[items_key]:
                        _item["partition"] = _pt_to_set
                    st.success(f"已将 {_pt_to_set if _pt_to_set else '无分区'} 应用到所有表。")
                    st.rerun()

        # 一键拉取按钮
        _pending_items = [x for x in st.session_state[items_key] if x.get("fetch_status") in ("pending", "error")]
        if _pending_items:
            _pending_count = len(_pending_items)
            _btn_label = f"一键拉取全部未拉取的表（{_pending_count} 张）" if _pending_count > 1 else f"一键拉取 {_pending_items[0]['name']}"
            if st.button(
                _btn_label,
                key=f"{state_prefix}_batch_fetch",
                type="primary",
                use_container_width=True
            ):
                _oe = get_odps_entry(
                    st.session_state.get("odps_ak", "").strip(),
                    st.session_state.get("odps_sk", "").strip(),
                    st.session_state.get("odps_project", "").strip(),
                    st.session_state.get("odps_endpoint", "").strip(),
                )
                _ok_count = 0
                _fail_count = 0
                _progress = st.progress(0.0)
                for _idx, _item in enumerate(_pending_items):
                    _progress.progress((_idx) / len(_pending_items))
                    with st.spinner(f"正在拉取 {_item['name']}（{_idx+1}/{len(_pending_items)}）..."):
                        try:
                            _schema = get_table_schema_text(_oe, _item["name"])
                            for saved_item in st.session_state[items_key]:
                                if saved_item["id"] == _item["id"]:
                                    saved_item["schema_text"] = _schema
                                    saved_item["fetch_status"] = "ok"
                                    saved_item["fetch_error"] = ""
                                    break
                            _ok_count += 1
                        except Exception as e:
                            for saved_item in st.session_state[items_key]:
                                if saved_item["id"] == _item["id"]:
                                    saved_item["fetch_status"] = "error"
                                    saved_item["fetch_error"] = str(e)
                                    break
                            _fail_count += 1
                _progress.progress(1.0)
                if _ok_count and not _fail_count:
                    st.success(f"全部拉取成功（{_ok_count} 张）。")
                elif _ok_count and _fail_count:
                    st.warning(f"成功 {_ok_count} 张，失败 {_fail_count} 张。失败的表可修正后重新拉取。")
                else:
                    st.error(f"全部拉取失败（{_fail_count} 张）。请检查表名和权限后重新拉取。")
                st.rerun()
        else:
            _has_items = bool(st.session_state[items_key])
            _all_ok = all(x.get("fetch_status") == "ok" for x in st.session_state[items_key]) if _has_items else False
            if _has_items and _all_ok:
                st.caption("所有表已拉取成功。")

    # ===== 统一渲染行列表 =====
    if not st.session_state[items_key]:
        st.info(f"暂无{title}。")
        return ""

    st.write(f"共 {len(st.session_state[items_key])} 张表：")

    for index, item in enumerate(list(st.session_state[items_key]), start=1):
        # 始终可见的一行：表名+状态 | 分区输入 | 删除
        _c_name, _c_pt, _c_del = st.columns([5, 3, 1])

        with _c_name:
            _status_tag = ""
            if item.get("source") == "odps":
                _status = item.get("fetch_status", "pending")
                if _status == "ok":
                    _status_tag = " ✅ 已拉取"
                elif _status == "error":
                    _status_tag = " ❌ 拉取失败"
                elif _status == "pending":
                    _status_tag = " ⏳ 待拉取"
            else:
                _status_tag = " 📄 xlsx"

            st.markdown(f"**{index}. {item['name']}**{_status_tag}")

            if item.get("source") == "odps" and item.get("fetch_status") == "error":
                st.error(f"拉取失败：{item.get('fetch_error', '未知错误')}")

        with _c_pt:
            _pt_val = st.text_input(
                "分区",
                value=item.get("partition", ""),
                key=f"{state_prefix}_pt_{item['id']}",
                placeholder="pt='20250101' 或留空=无分区",
                label_visibility="collapsed"
            )
            for saved_item in st.session_state[items_key]:
                if saved_item["id"] == item["id"]:
                    saved_item["partition"] = _pt_val.strip()
                    break

        with _c_del:
            if st.button("删除", key=f"{state_prefix}_del_{item['id']}"):
                st.session_state[items_key] = [
                    x for x in st.session_state[items_key]
                    if x["id"] != item["id"]
                ]
                st.session_state[uploader_version_key] += 1
                st.rerun()

        # 展开区：预览数据 + 表结构内容（默认收缩，不挡分区填写）
        _detail_label = f"展开详情：预览数据 / 表结构内容"
        with st.expander(_detail_label, expanded=False):

            # ODPS 模式：单行预览数据按钮
            if item.get("source") == "odps" and item.get("fetch_status") == "ok":
                if st.button("预览数据(20行)", key=f"{state_prefix}_pv_{item['id']}", use_container_width=True):
                    _oe = get_odps_entry(
                        st.session_state.get("odps_ak", "").strip(),
                        st.session_state.get("odps_sk", "").strip(),
                        st.session_state.get("odps_project", "").strip(),
                        st.session_state.get("odps_endpoint", "").strip(),
                    )
                    _pt_val = item.get("partition", "").strip()
                    with st.spinner(f"正在预览 {item['name']}..."):
                        try:
                            # 分区表没填分区条件时提前拦截，避免底层报错
                            # is_partitioned_table 返回 None 表示无法判断（权限不足等），不拦截
                            _is_pt = is_partitioned_table(_oe, item["name"])
                            if not _pt_val and _is_pt is True:
                                st.warning("这张是分区表，请在分区框里填入分区条件（例如 pt='20250101'）后再预览。")
                            else:
                                _df = preview_table_data(_oe, item["name"], _pt_val)
                                if _df.empty:
                                    st.info("该分区下没有数据。")
                                else:
                                    st.dataframe(_df, use_container_width=True, height=250)
                        except Exception as e:
                            st.error(f"预览失败：{e}")

            # 显示表结构内容（折叠）
            _show = st.checkbox(
                "显示表结构内容",
                key=f"{state_prefix}_show_{item['id']}"
            )
            if _show:
                _preview_text = item.get("schema_text", "")
                if not _preview_text:
                    if item.get("source") == "odps":
                        st.caption("暂无表结构内容，请先拉取。")
                    else:
                        st.caption("暂无内容。")
                else:
                    if len(_preview_text) > 8000:
                        _preview_text = _preview_text[:8000] + "\n\n......内容过长，已截断预览......"
                    st.text_area(
                        "表结构内容",
                        value=_preview_text,
                        height=260,
                        disabled=True,
                        key=f"{state_prefix}_preview_{item['id']}"
                    )

    # 拼接最终传给大模型的内容
    final_parts = []
    for item in st.session_state[items_key]:
        partition = item.get("partition", "").strip()
        partition_text = partition if partition else "无分区"
        schema_content = item.get("schema_text", "")
        if not schema_content:
            continue
        final_parts.append(
            f"""
### 表结构：{item['name']}

分区信息：{partition_text}

表结构内容：
{schema_content}
"""
        )
    return "\n\n".join(final_parts)
