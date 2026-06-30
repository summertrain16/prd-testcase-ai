"""
PRD 测试用例生成工具 — 主入口文件

文件结构：
- app.py            (本文件) 页面配置、session_state 初始化、侧边栏、4 步页面路由
- constants.py      全局常量（步骤名称）
- prompts.py        LLM Prompt 模板（4 个）
- llm_utils.py      LLM 调用（call_llm, is_llm_error）
- file_utils.py     文件读取（xlsx/pdf/docx/txt 等）
- markdown_utils.py Markdown 解析（待确认点、SQL 提取等）
- ui_components.py  UI 组件（CSS、步骤进度、表结构上传器等）
- odps_utils.py     PyODPS 连接与操作
"""

import os

import io

import pandas as pd
import streamlit as st
from dotenv import load_dotenv

from constants import STEP_INPUT, STEP_PENDING, STEP_FINAL, STEP_TEST_CASE
from prompts import (
    PRD_DRAFT_ANALYSIS_PROMPT,
    PRD_FINAL_ANALYSIS_PROMPT,
    PRD_ITERATIVE_PENDING_ANALYSIS_PROMPT,
    TEST_GEN_PROMPT,
)
from llm_utils import call_llm, is_llm_error
from file_utils import read_uploaded_file
from markdown_utils import (
    parse_pending_points_from_markdown,
    pending_points_to_llm_text,
    remove_pending_points_section,
    extract_sql_section_from_test_result,
    extract_sql_code_blocks,
    parse_markdown_tables,
)
from ui_components import (
    inject_custom_css,
    render_step_progress,
    render_page_header,
    render_markdown_in_scroll_box,
    render_test_case_result_with_download,
    render_pending_points_data_editor,
    sync_pending_points_from_widgets,
    render_table_schema_uploader,
    go_to_step,
    get_materials_from_state,
)
from odps_utils import get_odps_entry, run_single_sql, test_odps_connection

# =========================
# 1. 加载环境变量（仅作为默认值，实际从侧边栏 session_state 读取）
# =========================

load_dotenv()

_ENV_API_KEY = os.getenv("OPENAI_API_KEY", "")
_ENV_BASE_URL = os.getenv("OPENAI_BASE_URL", "https://api.openai.com/v1")
_ENV_MODEL = os.getenv("OPENAI_MODEL", "gpt-4o-mini")


# =========================
# 2. Streamlit 页面配置
# =========================

st.set_page_config(
    page_title="DataTest 自动化平台",
    page_icon=None,
    layout="wide"
)

inject_custom_css()

st.markdown(
    """
<div class="app-hero">
    <div class="app-hero-title">DataTest 自动化平台</div>
    <div class="app-hero-desc">
        智能解析 PRD 需求，自动生成数据测试用例与 SQL 校验脚本，驱动数据质量保障全流程。
    </div>
</div>
    """,
    unsafe_allow_html=True
)
render_step_progress()

# =========================
# 3. 初始化 Session State
# =========================

if "current_step" not in st.session_state:
    st.session_state["current_step"] = STEP_INPUT

if "prd_draft_analysis_result" not in st.session_state:
    st.session_state["prd_draft_analysis_result"] = ""

if "prd_pending_answers" not in st.session_state:
    st.session_state["prd_pending_answers"] = ""

if "prd_final_analysis_result" not in st.session_state:
    st.session_state["prd_final_analysis_result"] = ""

if "test_case_result" not in st.session_state:
    st.session_state["test_case_result"] = ""

if "uploaded_prd_text" not in st.session_state:
    st.session_state["uploaded_prd_text"] = ""

if "prd_file_uploader_version" not in st.session_state:
    st.session_state["prd_file_uploader_version"] = 0

if "source_schema_items" not in st.session_state:
    st.session_state["source_schema_items"] = []

if "result_schema_items" not in st.session_state:
    st.session_state["result_schema_items"] = []

if "source_schema_uploader_version" not in st.session_state:
    st.session_state["source_schema_uploader_version"] = 0

if "result_schema_uploader_version" not in st.session_state:
    st.session_state["result_schema_uploader_version"] = 0

if "pending_points_rows" not in st.session_state:
    st.session_state["pending_points_rows"] = []
if "pending_points_editor_version" not in st.session_state:
    st.session_state["pending_points_editor_version"] = 0

if "prd_current_analysis_result" not in st.session_state:
    st.session_state["prd_current_analysis_result"] = ""

if "pending_analysis_round" not in st.session_state:
    st.session_state["pending_analysis_round"] = 0

if "pending_confirm_history" not in st.session_state:
    st.session_state["pending_confirm_history"] = ""

if "ignore_remaining_pending_points" not in st.session_state:
    st.session_state["ignore_remaining_pending_points"] = False

if "ignored_pending_points_text" not in st.session_state:
    st.session_state["ignored_pending_points_text"] = ""

material_state_defaults = {
    "prd_text": "",
    "prd_manual_text": "",
    "meeting_notes": "",
    "result_table_schema": "",
    "source_table_schema": "",
    "dev_code": "",
}
for key, default_value in material_state_defaults.items():
    if key not in st.session_state:
        st.session_state[key] = default_value


# =========================
# 4. 页面侧边栏
# =========================

with st.sidebar:
    # 模型配置折叠起来
    with st.expander("模型配置", expanded=False):
        st.caption("每个用户填写自己的 API Key，互不影响。配置仅保存在浏览器当前会话中。")

        # 初始化 session_state（用 .env 作为默认值）
        if "user_openai_api_key" not in st.session_state:
            st.session_state["user_openai_api_key"] = _ENV_API_KEY
        if "user_openai_base_url" not in st.session_state:
            st.session_state["user_openai_base_url"] = _ENV_BASE_URL
        if "user_openai_model" not in st.session_state:
            st.session_state["user_openai_model"] = _ENV_MODEL

        st.text_input(
            "API Key",
            key="user_openai_api_key",
            type="password",
            placeholder="sk-xxxx",
            help="你的 LLM API Key，例如 sk-xxx"
        )
        st.text_input(
            "Base URL",
            key="user_openai_base_url",
            placeholder="http://xxx/v1",
            help="LLM 接口地址，注意以 /v1 结尾"
        )
        st.text_input(
            "Model",
            key="user_openai_model",
            placeholder="gpt-4o-mini",
            help="模型名称"
        )

        configured = bool(st.session_state.get("user_openai_api_key", "").strip())
        if configured:
            st.success("API Key 已配置")
        else:
            st.warning("请填写 API Key")

    # ===== ODPS 连接配置 =====
    with st.expander("ODPS 连接配置", expanded=False):
        st.caption("连接 MaxCompute/DataWorks。配置一次后同一会话内不丢失。也可在 Streamlit Cloud Settings → Secrets 里持久化。")

        # 从 st.secrets 读默认值（如果配了的话）
        _secrets_endpoint = st.secrets.get("odps_endpoint", "") if hasattr(st, 'secrets') else ""
        _secrets_project = st.secrets.get("odps_project", "") if hasattr(st, 'secrets') else ""
        _secrets_ak = st.secrets.get("odps_ak", "") if hasattr(st, 'secrets') else ""
        _secrets_sk = st.secrets.get("odps_sk", "") if hasattr(st, 'secrets') else ""

        if "odps_endpoint" not in st.session_state:
            st.session_state["odps_endpoint"] = _secrets_endpoint
        if "odps_project" not in st.session_state:
            st.session_state["odps_project"] = _secrets_project
        if "odps_ak" not in st.session_state:
            st.session_state["odps_ak"] = _secrets_ak
        if "odps_sk" not in st.session_state:
            st.session_state["odps_sk"] = _secrets_sk

        st.text_input("Endpoint", key="odps_endpoint",
            placeholder="http://service.odps.aliyun.com/api",
            help="MaxCompute 访问地址。公网一般填 http://service.odps.aliyun.com/api")
        st.text_input("Project（可选，默认项目）", key="odps_project",
            placeholder="例如：ods_project。留空也行，表名带前缀即可",
            help="默认 MaxCompute 项目名。留空时表名要写 项目名.表名 跨项目查")
        st.text_input("AccessKey ID", key="odps_ak", type="password")
        st.text_input("AccessKey Secret", key="odps_sk", type="password")

        _c_test, _c_status = st.columns([1, 2])
        with _c_test:
            if st.button("测试连接", key="test_odps_btn", use_container_width=True):
                _test_entry = get_odps_entry(
                    st.session_state.get("odps_ak", "").strip(),
                    st.session_state.get("odps_sk", "").strip(),
                    st.session_state.get("odps_project", "").strip(),
                    st.session_state.get("odps_endpoint", "").strip(),
                )
                if _test_entry is None:
                    st.error("请先填写 Endpoint、AccessKey ID、AccessKey Secret")
                else:
                    with st.spinner("测试中..."):
                        _ok, _msg = test_odps_connection(_test_entry)
                    if _ok:
                        st.success(_msg)
                    else:
                        st.error(f"连接失败：{_msg}")
        with _c_status:
            _has_ak = bool(st.session_state.get("odps_ak", "").strip())
            if _has_ak:
                st.caption("已填写配置。如需持久化，见下方说明。")
            else:
                st.info("未配置 ODPS，仍可用 xlsx 上传")

        # 持久化提示
        if not _secrets_ak:
            with st.expander("如何让配置不丢失？", expanded=False):
                st.markdown("""
在 Streamlit Cloud 管理台 → 你的 app → Settings → Secrets 中添加：

```toml
odps_endpoint = "http://service.odps.aliyun.com/api"
odps_project = "你的项目名"
odps_ak = "你的AccessKey ID"
odps_sk = "你的AccessKey Secret"
```

保存后重启 app，配置会自动填入，不用每次手填。
本地运行则在项目根目录建 `.streamlit/secrets.toml`，格式同上。
""")

    st.divider()
    if st.button("清空全部结果", use_container_width=True):
        st.session_state["prd_draft_analysis_result"] = ""
        st.session_state["prd_pending_answers"] = ""
        st.session_state["prd_final_analysis_result"] = ""
        st.session_state["test_case_result"] = ""
        st.session_state["uploaded_prd_text"] = ""
        st.session_state["prd_file_uploader_version"] += 1
        st.session_state["source_schema_items"] = []
        st.session_state["result_schema_items"] = []
        st.session_state["source_schema_uploader_version"] += 1
        st.session_state["result_schema_uploader_version"] += 1
        st.session_state["pending_points_rows"] = []
        st.session_state["pending_points_editor_version"] += 1
        st.session_state["prd_current_analysis_result"] = ""
        st.session_state["pending_analysis_round"] = 0
        st.session_state["pending_confirm_history"] = ""
        st.session_state["ignore_remaining_pending_points"] = False
        st.session_state["ignored_pending_points_text"] = ""
        st.session_state["prd_text"] = ""
        st.session_state["prd_manual_text"] = ""
        st.session_state["meeting_notes"] = ""
        st.session_state["result_table_schema"] = ""
        st.session_state["source_table_schema"] = ""
        st.session_state["dev_code"] = ""
        st.session_state["current_step"] = STEP_INPUT
        # 清空 ODPS 表结构相关 state
        st.session_state["source_schema_mode"] = ""
        st.session_state["result_schema_mode"] = ""
        # 清空 SQL 执行结果缓存
        st.session_state["sql_run_results"] = {}
        # 清空 ODPS 连接缓存（下次用新配置重建）
        get_odps_entry.clear()
        st.rerun()


# =========================
# 5. 第 1 步：PRD 输入区 + 补充信息区
# =========================

if st.session_state["current_step"] == STEP_INPUT:
    render_page_header(
        title="1. 输入材料",
        desc="上传或粘贴 PRD，并补充会议纪要、源表表结构、结果表表结构、分区信息和开发代码。",
        icon=""
    )

    # =========================
    # PRD 输入区
    # =========================

    st.header("一、输入 PRD")

    with st.expander("上传 PRD 文件", expanded=True):
        prd_file = st.file_uploader(
            "支持 txt、md、pdf、docx、xlsx、sql、csv、json、py",
            type=["txt", "md", "pdf", "docx", "xlsx", "sql", "csv", "json", "py"],
            key=f"prd_file_{st.session_state['prd_file_uploader_version']}"
        )

        if prd_file is not None:
            uploaded_text = read_uploaded_file(prd_file)
            st.session_state["uploaded_prd_text"] = uploaded_text

        if st.session_state.get("uploaded_prd_text", ""):
            st.success("已读取上传的 PRD 文件。")

            if st.button("清除已上传内容"):
                st.session_state["uploaded_prd_text"] = ""
                st.session_state["prd_file_uploader_version"] += 1
                st.rerun()
        else:
            st.info("未上传 PRD 文件，可在下方粘贴内容。")

    with st.expander("粘贴 PRD 内容", expanded=True):
        prd_manual_text = st.text_area(
            "粘贴 PRD 内容",
            key="prd_manual_text",
            height=260,
            placeholder="请在这里粘贴 PRD 文本。如果已经上传文件，也可以在这里补充说明。",
            label_visibility="collapsed"
        )

    prd_text = ""

    if st.session_state.get("uploaded_prd_text", ""):
        prd_text += st.session_state["uploaded_prd_text"]

    if prd_manual_text.strip():
        prd_text += "\n\n【用户手动补充 PRD 内容】\n" + prd_manual_text

    st.session_state["prd_text"] = prd_text

    if prd_text.strip():
        with st.expander("查看当前 PRD 原文", expanded=False):
            st.text_area(
                "当前合并后的 PRD 内容",
                value=prd_text,
                height=300,
                disabled=True
            )

    # =========================
    # 补充信息区
    # =========================

    st.header("二、补充信息，可选")

    with st.expander("填写会议纪要、表结构、分区、开发代码等补充信息", expanded=False):
        meeting_notes = st.text_area(
            "会议纪要，可选",
            key="meeting_notes",
            height=140,
            placeholder="例如：会议中确认了统计口径、过滤条件、字段含义等。"
        )


        st.divider()

        result_table_schema = render_table_schema_uploader(
            title="结果表表结构",
            state_prefix="result_schema"
        )

        st.session_state["result_table_schema"] = result_table_schema

        st.divider()

        source_table_schema = render_table_schema_uploader(
            title="源表表结构",
            state_prefix="source_schema"
        )

        st.session_state["source_table_schema"] = source_table_schema

        st.divider()

        dev_code = st.text_area(
            "参考开发代码，可选",
            key="dev_code",
            height=220,
            placeholder="可以粘贴 SQL、PySpark、DataWorks 调度代码等。"
        )

    # =========================
    # 第一步：生成初版需求提炼和待确认点
    # =========================

    st.header("三、第一步：生成初版需求提炼和待确认点")

    st.info(
        "第一步会根据 PRD、会议纪要、表结构、分区信息、开发代码进行初版分析；不确定的内容会放到待确认点中。"
    )

    _generate_draft_clicked = st.button(
        "🚀 生成初版需求提炼和待确认点",
        type="primary"
    )

    if _generate_draft_clicked:
        materials = get_materials_from_state()

        if not materials["prd_text"].strip():
            st.warning("请先上传或粘贴 PRD 内容。")
        else:
            with st.spinner("正在分析 PRD..."):
                user_content = f"""
以下是 PRD 原文：

{materials["prd_text"]}

以下是会议纪要：

{materials["meeting_notes"] if materials["meeting_notes"].strip() else "未提供"}

以下是结果表表结构及分区信息：

{materials["result_table_schema"] if materials["result_table_schema"].strip() else "未上传结果表表结构。"}

以下是源表表结构及分区信息：

{materials["source_table_schema"] if materials["source_table_schema"].strip() else "未上传源表表结构。"}

以下是开发代码：

{materials["dev_code"] if materials["dev_code"].strip() else "未提供"}

请基于以上全部信息进行第一轮分析。

请特别注意：
1. 测试用例需要精简，但不能把所有字段的一致性比对强行混在一个用例里。
2. 最重要的测试是"主键唯一性校验"和"结果表与源表加工结果一致性比对"。
3. 字段一致性比对需要按来源表、加工逻辑、过滤条件、关联条件、分区条件和复杂度合理拆分。
4. 简单同源、同关联、同过滤、同分区、直接映射的字段，可以合并到同一条一致性比对用例。
5. 复杂字段、金额字段、枚举映射字段、case when 字段、聚合字段、去重字段、不同源表字段，应单独生成一致性比对用例和 SQL。
6. 每条一致性比对 SQL 的 final select 必须同时展示源表加工后的字段值和结果表字段值。
7. 如果存在差异，SQL 结果应能直观看到 expected/source 值、actual/result 值、diff_flag 和 diff_type。
8. 不要只输出差异数量，要输出差异明细。
9. 复杂字段的 expected CTE 中要保留必要中间字段，方便排查差异原因。
"""

                draft_result = call_llm(
                    PRD_DRAFT_ANALYSIS_PROMPT,
                    user_content
                )
                if is_llm_error(draft_result):
                    st.error(draft_result)
                    st.stop()

                st.session_state["prd_draft_analysis_result"] = draft_result
                st.session_state["prd_current_analysis_result"] = draft_result

                st.session_state["pending_points_rows"] = parse_pending_points_from_markdown(
                    draft_result
                )

                st.session_state["pending_points_editor_version"] += 1

                st.session_state["prd_pending_answers"] = pending_points_to_llm_text(
                    st.session_state["pending_points_rows"]
                )

                st.session_state["pending_analysis_round"] = 1
                st.session_state["pending_confirm_history"] = ""

                st.session_state["ignore_remaining_pending_points"] = False
                st.session_state["ignored_pending_points_text"] = ""

                st.session_state["prd_final_analysis_result"] = ""
                st.session_state["test_case_result"] = ""

            go_to_step(STEP_PENDING)


# =========================
# 6. 第 2 步：展示当前需求分析结果 + 待确认点多轮收敛
# =========================

elif st.session_state["current_step"] == STEP_PENDING:
    render_page_header(
        title="2. 待确认点收敛",
        desc="查看 AI 初版需求分析结果，并在待确认点清单中补充说明，支持多轮收敛。",
        icon=""
    )

    if not st.session_state.get("prd_current_analysis_result"):
        st.warning("请先完成第 1 步：输入材料并生成初版需求分析。")

        if st.button("返回第 1 步"):
            go_to_step(STEP_INPUT)

        st.stop()

    materials = get_materials_from_state()

    st.subheader("当前需求分析结果")

    current_without_pending_points = remove_pending_points_section(
        st.session_state["prd_current_analysis_result"]
    )

    # ===== 下载 PRD 分析结果（Excel） =====
    _download_excel_data = io.BytesIO()
    with pd.ExcelWriter(_download_excel_data, engine="openpyxl") as _writer:
        # Sheet1: 需求分析结果
        _analysis_text = current_without_pending_points.strip() if current_without_pending_points.strip() else st.session_state["prd_current_analysis_result"]
        _analysis_df = pd.DataFrame({"需求分析结果": _analysis_text.split("\n")})
        _analysis_df.to_excel(_writer, sheet_name="需求分析结果", index=False)

        # Sheet2: 待确认点清单
        _pending_rows = st.session_state["pending_points_rows"]
        if _pending_rows:
            _pending_df = pd.DataFrame(_pending_rows)
        else:
            _pending_df = pd.DataFrame({"说明": ["无待确认点"]})
        _pending_df.to_excel(_writer, sheet_name="待确认点清单", index=False)

        # Sheet3: 历轮处理记录
        _history_text = st.session_state.get("pending_confirm_history", "").strip()
        if _history_text:
            _history_df = pd.DataFrame({"历轮待确认点处理记录": _history_text.split("\n")})
        else:
            _history_df = pd.DataFrame({"说明": ["暂无历史记录"]})
        _history_df.to_excel(_writer, sheet_name="历轮处理记录", index=False)

    _download_excel_data.seek(0)

    st.download_button(
        label="📥 下载 PRD 分析结果（Excel）",
        data=_download_excel_data,
        file_name="prd_当前需求分析和待确认点.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.caption(
        f"当前待确认点解析轮次：第 {st.session_state.get('pending_analysis_round', 1)} 轮"
    )

    if st.session_state["pending_points_rows"]:
        if current_without_pending_points.strip():
            render_markdown_in_scroll_box(
                title="查看当前需求分析结果",
                markdown_text=current_without_pending_points,
                height=600,
                expanded=True
            )
    else:
        render_markdown_in_scroll_box(
            title="查看当前完整需求分析结果",
            markdown_text=st.session_state["prd_current_analysis_result"],
            height=600,
            expanded=True
        )

    st.subheader("待确认点清单")

    st.caption(
        "请直接在最后一列『用户补充说明』中填写确认结果。填写后可点击『提交补充说明，继续解析待确认点』。"
    )

    if not st.session_state["pending_points_rows"]:
        st.success("当前没有阻塞测试设计或 SQL 校验的待确认点，可以继续生成最终版需求提炼表。")
        st.session_state["prd_pending_answers"] = "无待确认点。"

        if st.button(
            "进入第 3 步：生成最终版",
            type="primary"
        ):
                go_to_step(STEP_FINAL)

    else:
        render_pending_points_data_editor()

        st.divider()

        col_a, col_b = st.columns([1, 1])

        with col_a:
            continue_pending_clicked = st.button(
                "🔄 提交补充说明，继续收敛",
                type="primary",
                use_container_width=True
            )

        with col_b:
            ignore_pending_clicked = st.button(
                "⏭️ 忽略剩余待确认点，继续",
                use_container_width=True
            )

        if continue_pending_clicked:
            sync_pending_points_from_widgets()

            with st.spinner("正在重新解析..."):
                current_answers = st.session_state["prd_pending_answers"]

                user_content = f"""
以下是 PRD 原文：

{materials["prd_text"]}

以下是上一轮 AI 需求分析结果：

{st.session_state["prd_current_analysis_result"]}

以下是本轮用户针对待确认点填写的补充说明：

{current_answers if current_answers.strip() else "用户未补充。"}

以下是历轮待确认点处理记录：

{st.session_state["pending_confirm_history"] if st.session_state["pending_confirm_history"].strip() else "暂无历史记录。"}

以下是会议纪要：

{materials["meeting_notes"] if materials["meeting_notes"].strip() else "未提供"}

以下是结果表表结构及分区信息：

{materials["result_table_schema"] if materials["result_table_schema"].strip() else "未上传结果表表结构。"}

以下是源表表结构及分区信息：

{materials["source_table_schema"] if materials["source_table_schema"].strip() else "未上传源表表结构。"}

以下是开发代码：

{materials["dev_code"] if materials["dev_code"].strip() else "未提供"}

请基于以上信息进行新一轮需求解析，并重新输出仍需确认的待确认点清单。
"""

                new_analysis_result = call_llm(
                    PRD_ITERATIVE_PENDING_ANALYSIS_PROMPT,
                    user_content
                )
                if is_llm_error(new_analysis_result):
                    st.error(new_analysis_result)
                    st.stop()

                st.session_state["pending_confirm_history"] += f"""

==============================
第 {st.session_state["pending_analysis_round"]} 轮待确认点处理记录
==============================

【用户补充说明】
{current_answers if current_answers.strip() else "用户未补充。"}

【本轮解析前分析结果】
{st.session_state["prd_current_analysis_result"]}

【本轮解析后结果】
{new_analysis_result}
"""

                st.session_state["prd_current_analysis_result"] = new_analysis_result
                st.session_state["prd_draft_analysis_result"] = new_analysis_result

                st.session_state["pending_points_rows"] = parse_pending_points_from_markdown(
                    new_analysis_result
                )

                st.session_state["pending_points_editor_version"] += 1

                st.session_state["prd_pending_answers"] = pending_points_to_llm_text(
                    st.session_state["pending_points_rows"]
                )

                st.session_state["pending_analysis_round"] += 1

                st.session_state["ignore_remaining_pending_points"] = False
                st.session_state["ignored_pending_points_text"] = ""

                st.session_state["prd_final_analysis_result"] = ""
                st.session_state["test_case_result"] = ""

            st.rerun()

        if ignore_pending_clicked:
            sync_pending_points_from_widgets()

            st.session_state["ignore_remaining_pending_points"] = True

            st.session_state["ignored_pending_points_text"] = pending_points_to_llm_text(
                st.session_state["pending_points_rows"]
            )

            st.success("已标记忽略剩余待确认点，可以继续生成最终版需求提炼表。")

            go_to_step(STEP_FINAL)

    if st.session_state.get("ignore_remaining_pending_points", False):
        st.warning(
            "你已选择忽略剩余待确认点。后续生成最终版和测试用例时，AI 会基于当前信息继续处理；被忽略的问题可能在最终版备注或 SQL TODO 中体现。"
        )


# =========================
# 7. 第 3 步：生成并展示最终版需求提炼表
# =========================

elif st.session_state["current_step"] == STEP_FINAL:
    render_page_header(
        title="3. 最终版需求提炼",
        desc="基于 PRD、补充说明、表结构、分区信息和待确认点处理记录，生成最终版需求提炼表。",
        icon=""
    )

    if not st.session_state.get("prd_current_analysis_result"):
        st.warning("请先完成第 1 步：输入材料并生成初版需求分析。")

        if st.button("返回第 1 步"):
            go_to_step(STEP_INPUT)

        st.stop()

    has_pending_points = bool(st.session_state.get("pending_points_rows", []))
    ignored_pending = st.session_state.get("ignore_remaining_pending_points", False)

    if has_pending_points and not ignored_pending:
        st.warning("当前仍存在待确认点。请先在第 2 步补充说明，或者选择忽略剩余待确认点。")

        if st.button("返回第 2 步处理待确认点"):
            go_to_step(STEP_PENDING)

        st.stop()

    materials = get_materials_from_state()

    if ignored_pending:
        st.info("当前存在被用户选择忽略的待确认点，仍允许生成最终版需求提炼表。")
    else:
        st.success("当前无待确认点，可以生成最终版需求提炼表。")

    _gen_final_clicked = st.button(
        "生成最终版需求提炼表",
        type="primary"
    )

    if _gen_final_clicked:
        with st.spinner("正在生成最终版..."):
            sync_pending_points_from_widgets()

            user_content = f"""
以下是 PRD 原文：

{materials["prd_text"]}

以下是经过多轮待确认点收敛后的最新需求分析结果：

{st.session_state["prd_current_analysis_result"]}

以下是历轮待确认点处理记录：

{st.session_state["pending_confirm_history"] if st.session_state["pending_confirm_history"].strip() else "暂无历史记录。"}

以下是当前待确认点及用户补充说明：

{st.session_state["prd_pending_answers"] if st.session_state["prd_pending_answers"].strip() else "无待确认点。"}

以下是用户选择忽略的剩余待确认点：

{st.session_state["ignored_pending_points_text"] if st.session_state.get("ignore_remaining_pending_points", False) else "无。"}

以下是会议纪要：

{materials["meeting_notes"] if materials["meeting_notes"].strip() else "未提供"}

以下是结果表表结构及分区信息：

{materials["result_table_schema"] if materials["result_table_schema"].strip() else "未上传结果表表结构。"}

以下是源表表结构及分区信息：

{materials["source_table_schema"] if materials["source_table_schema"].strip() else "未上传源表表结构。"}

以下是开发代码：

{materials["dev_code"] if materials["dev_code"].strip() else "未提供"}

请基于以上信息生成最终版需求提炼表。

特别要求：
1. 对已经确认清楚的问题，状态标记为『已明确』。
2. 对用户明确选择忽略的待确认点，不要继续阻塞最终版输出。
3. 如果用户忽略的问题会影响 SQL 准确性，需要在测试关注点、备注或仍需确认问题中标记『用户选择忽略』。
4. 不要再次输出低价值或不影响测试 SQL 的待确认问题。
5. 最终版需求提炼必须能直接用于后续生成数据测试用例和 SQL 校验脚本。
"""

            final_result = call_llm(
                PRD_FINAL_ANALYSIS_PROMPT,
                user_content
            )
            if is_llm_error(final_result):
                st.error(final_result)
                st.stop()

            st.session_state["prd_final_analysis_result"] = final_result
            st.session_state["test_case_result"] = ""

        st.success("最终版需求提炼表已生成。")

    if st.session_state["prd_final_analysis_result"]:
        st.subheader("最终版需求提炼表")

        render_markdown_in_scroll_box(
            title="查看最终版需求提炼表",
            markdown_text=st.session_state["prd_final_analysis_result"],
            height=700,
            expanded=True
        )

        # ===== 下载最终版需求提炼表（Excel） =====
        _final_text = st.session_state["prd_final_analysis_result"]
        _final_tables = parse_markdown_tables(_final_text)

        _final_excel_data = io.BytesIO()
        with pd.ExcelWriter(_final_excel_data, engine="openpyxl") as _fwriter:
            if _final_tables:
                for _fti, _ftable in enumerate(_final_tables, 1):
                    # Sheet 名最长 31 字符，Excel 限制
                    _sheet_name = _ftable["title"][:31] if _ftable["title"] else f"Table{_fti}"
                    _final_df = pd.DataFrame(_ftable["rows"], columns=_ftable["columns"])
                    _final_df.to_excel(_fwriter, sheet_name=_sheet_name, index=False)
            else:
                # 没有表格，整段文本放一个 Sheet
                _final_df = pd.DataFrame({"需求提炼表": _final_text.split("\n")})
                _final_df.to_excel(_fwriter, sheet_name="需求提炼表", index=False)

        _final_excel_data.seek(0)

        st.download_button(
            label="📥 下载需求提炼表（Excel）",
            data=_final_excel_data,
            file_name="prd_最终版需求提炼表.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.divider()

        if st.button(
            "进入第 4 步：生成测试用例",
            type="primary"
        ):
                go_to_step(STEP_TEST_CASE)


# =========================
# 8. 第 4 步：生成并展示测试用例和 SQL
# =========================

elif st.session_state["current_step"] == STEP_TEST_CASE:
    render_page_header(
        title="4. 测试用例与 SQL",
        desc="根据最终版需求提炼表生成精简、可定位问题的数据测试用例和 SQL 校验脚本。",
        icon=""
    )

    if not st.session_state.get("prd_final_analysis_result"):
        st.warning("请先完成第 3 步：生成最终版需求提炼表。")

        if st.button("返回第 3 步"):
            go_to_step(STEP_FINAL)

        st.stop()

    materials = get_materials_from_state()

    _gen_test_clicked = st.button(
        "生成测试用例和 SQL",
        type="primary"
    )

    if _gen_test_clicked:
        with st.spinner("正在生成测试用例..."):
            sync_pending_points_from_widgets()

            user_content = f"""
以下是 PRD 原文：

{materials["prd_text"]}

以下是最终版需求提炼表：

{st.session_state["prd_final_analysis_result"]}

以下是历轮待确认点处理记录：

{st.session_state["pending_confirm_history"] if st.session_state["pending_confirm_history"].strip() else "暂无历史记录。"}

以下是当前待确认点及用户补充说明：

{st.session_state["prd_pending_answers"] if st.session_state["prd_pending_answers"].strip() else "无待确认点。"}

以下是用户选择忽略的剩余待确认点：

{st.session_state["ignored_pending_points_text"] if st.session_state.get("ignore_remaining_pending_points", False) else "无。"}

以下是补充信息。补充信息均为可选，如果没有提供，请不要自行脑补。

【会议纪要】
{materials["meeting_notes"] if materials["meeting_notes"].strip() else "未提供"}

【结果表表结构 / 字段说明 / 分区信息】
{materials["result_table_schema"] if materials["result_table_schema"].strip() else "未上传结果表表结构。"}

【源表表结构 / 字段说明 / 分区信息】
{materials["source_table_schema"] if materials["source_table_schema"].strip() else "未上传源表表结构。"}

【参考开发代码】
{materials["dev_code"] if materials["dev_code"].strip() else "未提供"}

请基于以上信息生成数据测试用例和 SQL 校验脚本。

要求：
1. 必须以『最终版需求提炼表』为主要依据。
2. 如果最终版需求提炼表中仍存在『待确认』『部分待确认』『存在冲突』的内容，需要在测试用例中标记。
3. 如果存在『用户选择忽略』的待确认点，需要在测试用例前置条件、SQL 注释或待确认问题清单中标记。
4. 不允许根据初版分析中的不确定内容自行脑补。
5. 如果源表或结果表分区未提供，请在 SQL 中使用【分区字段】或 pt='YYYYMMDD' 占位，并在待确认问题中说明。
6. 如果某个上传表结构的分区信息为『无分区』，不要强行给该表添加分区条件。
"""

            test_result = call_llm(
                TEST_GEN_PROMPT,
                user_content
            )
            if is_llm_error(test_result):
                st.error(test_result)
                st.stop()

            st.session_state["test_case_result"] = test_result

        st.success("测试用例和 SQL 校验脚本已生成。")

    if st.session_state["test_case_result"]:
        st.subheader("测试用例和 SQL 校验脚本")

        render_test_case_result_with_download(
            st.session_state["test_case_result"]
        )

        # ===== 在线执行校验 SQL =====
        _odps_entry_run = get_odps_entry(
            st.session_state.get("odps_ak", "").strip(),
            st.session_state.get("odps_sk", "").strip(),
            st.session_state.get("odps_project", "").strip(),
            st.session_state.get("odps_endpoint", "").strip(),
        )

        if _odps_entry_run:
            st.divider()
            st.subheader("在线执行校验 SQL")

            _sql_section = extract_sql_section_from_test_result(
                st.session_state["test_case_result"]
            )
            _sql_blocks = extract_sql_code_blocks(_sql_section)

            if not _sql_blocks:
                st.info("未提取到可执行的 SQL 代码块。")
            else:
                st.caption(f"共提取到 {len(_sql_blocks)} 段 SQL，可逐段或一键执行。")

                # 初始化执行结果存储
                if "sql_run_results" not in st.session_state:
                    st.session_state["sql_run_results"] = {}

                # 一键执行全部
                _col_batch, _col_clear = st.columns([3, 1])
                with _col_batch:
                    if st.button("一键执行全部 SQL", type="primary", use_container_width=True, key="batch_run_sql"):
                        _progress = st.progress(0.0)
                        for _i, _sql_block in enumerate(_sql_blocks, 1):
                            _progress.progress((_i - 1) / len(_sql_blocks))
                            with st.spinner(f"执行 SQL-{_i:03d}（{_i}/{len(_sql_blocks)}）..."):
                                _df_r, _err_r = run_single_sql(_odps_entry_run, _sql_block)
                            st.session_state["sql_run_results"][_i] = {"df": _df_r, "err": _err_r}
                        _progress.progress(1.0)
                        _ok = sum(1 for v in st.session_state["sql_run_results"].values() if not v["err"])
                        _fail = len(st.session_state["sql_run_results"]) - _ok
                        if _fail == 0:
                            st.success(f"全部执行完成（{_ok} 段）。")
                        else:
                            st.warning(f"执行完成：成功 {_ok} 段，失败 {_fail} 段。")
                        st.rerun()
                with _col_clear:
                    if st.button("清空执行结果", use_container_width=True, key="clear_run_results"):
                        st.session_state["sql_run_results"] = {}
                        st.rerun()

                # 逐段展示 + 执行
                for _i, _sql_block in enumerate(_sql_blocks, 1):
                    with st.expander(f"SQL-{_i:03d}", expanded=False):
                        st.code(_sql_block, language="sql")

                        _c_run, _c_export = st.columns([3, 2])
                        with _c_run:
                            if st.button("执行", key=f"run_sql_{_i}"):
                                with st.spinner("执行中..."):
                                    _df_result, _err = run_single_sql(_odps_entry_run, _sql_block)
                                st.session_state["sql_run_results"][_i] = {"df": _df_result, "err": _err}
                                st.rerun()

                        # 展示已有结果
                        _cached = st.session_state["sql_run_results"].get(_i)
                        if _cached:
                            if _cached["err"]:
                                st.error(f"执行失败：{_cached['err']}")
                            elif _cached["df"].empty:
                                st.success("校验通过，无差异")
                            else:
                                st.caption(f"返回 {len(_cached['df'])} 行")
                                st.dataframe(
                                    _cached["df"],
                                    use_container_width=True,
                                    height=300
                                )
                                with _c_export:
                                    _csv_data = _cached["df"].to_csv(index=False).encode("utf-8-sig")
                                    st.download_button(
                                        label="导出 CSV",
                                        data=_csv_data,
                                        file_name=f"sql_{_i:03d}_result.csv",
                                        mime="text/csv",
                                        use_container_width=True,
                                        key=f"download_csv_{_i}"
                                    )
        else:
            st.info("未配置 ODPS 连接，仅支持下载 SQL 脚本手动执行。配置方法见侧边栏 ODPS 连接配置。")
else:
    st.warning("当前步骤状态异常，已返回第 1 步。")
    st.session_state["current_step"] = STEP_INPUT
    st.rerun()
