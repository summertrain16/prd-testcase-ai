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
from datetime import datetime

import pandas as pd
import streamlit as st
from dotenv import load_dotenv

from constants import STEP_INPUT, STEP_PENDING, STEP_FINAL, STEP_TEST_CASE
from prompts import (
    PRD_DRAFT_ANALYSIS_PROMPT,
    PRD_FINAL_ANALYSIS_PROMPT,
    PRD_ITERATIVE_PENDING_ANALYSIS_PROMPT,
    TEST_GEN_PROMPT,
    SQL_DIFF_ANALYSIS_PROMPT,
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
    render_empty_state,
    render_error_with_fold,
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


def _ts_filename(base: str, ext: str) -> str:
    """生成带时间戳的文件名，避免多次下载覆盖。"""
    _ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"{base}_{_ts}.{ext}"


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
<div class="app-hero" role="banner">
    <h1 class="app-hero-title">DataTest 自动化平台</h1>
    <p class="app-hero-desc">
        智能解析 PRD 需求，自动生成数据测试用例与 SQL 校验脚本，驱动数据质量保障全流程。
    </p>
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
        st.session_state["batch_diff_analysis"] = ""
        st.session_state["_sql_batch_running"] = False
        st.session_state["_sql_batch_idx"] = 0
        # 清空 ODPS 连接缓存（下次用新配置重建）
        get_odps_entry.clear()
        st.rerun()


# =========================
# 5. 第 1 步：PRD 输入区 + 补充信息区
# =========================

if st.session_state["current_step"] == STEP_INPUT:
    render_page_header(
        title="1. 输入材料",
        desc="上传或粘贴 PRD，补充表结构和参考信息，一键生成初版需求提炼。",
        icon=""
    )

    # =========================
    # Tab 布局：PRD / 表结构 / 补充说明 / 开发代码
    # =========================

    _tab_prd, _tab_schema, _tab_notes, _tab_code = st.tabs([
        "📄 PRD",
        "🗄️ 表结构",
        "📝 补充说明",
        "💻 开发代码",
    ])

    # ----- Tab 1: PRD -----
    with _tab_prd:
        st.caption("上传 PRD 文件或直接粘贴文本，二选一即可。")
        _col_prd_left, _col_prd_right = st.columns(2)

        with _col_prd_left:
            prd_file = st.file_uploader(
                "📎 上传 PRD 文件（txt、md、pdf、docx、xlsx、sql、csv、json、py）",
                type=["txt", "md", "pdf", "docx", "xlsx", "sql", "csv", "json", "py"],
                key=f"prd_file_{st.session_state['prd_file_uploader_version']}",
            )

            if prd_file is not None:
                uploaded_text = read_uploaded_file(prd_file)
                st.session_state["uploaded_prd_text"] = uploaded_text

            if st.session_state.get("uploaded_prd_text", ""):
                st.success(f"✅ 已读取：{prd_file.name if prd_file else '上传的文件'}")

                if st.button("🗑️ 清除已上传文件"):
                    st.session_state["uploaded_prd_text"] = ""
                    st.session_state["prd_file_uploader_version"] += 1
                    st.rerun()
            else:
                st.caption("未上传文件，可在右侧粘贴内容。")

        with _col_prd_right:
            prd_manual_text = st.text_area(
                "✏️ 粘贴 PRD 内容",
                key="prd_manual_text",
                height=260,
                placeholder="请在这里粘贴 PRD 文本。\n如果已经上传文件，也可以在这里补充说明。",
            )

    # ----- Tab 2: 表结构 -----
    with _tab_schema:
        st.caption("从 ODPS 在线拉取结果表和源表结构，填写分区条件后即可用于分析。")
        _col_schema_left, _col_schema_right = st.columns(2)

        with _col_schema_left:
            result_table_schema = render_table_schema_uploader(
                title="结果表表结构",
                state_prefix="result_schema"
            )
            st.session_state["result_table_schema"] = result_table_schema

        with _col_schema_right:
            source_table_schema = render_table_schema_uploader(
                title="源表表结构",
                state_prefix="source_schema"
            )
            st.session_state["source_table_schema"] = source_table_schema

    # ----- Tab 3: 补充说明 -----
    with _tab_notes:
        st.caption("粘贴会议纪要、评审记录等，帮助 AI 更准确理解需求口径。")
        meeting_notes = st.text_area(
            "会议纪要 / 评审记录",
            key="meeting_notes",
            height=220,
            placeholder="例如：\n- 会议中确认了订单状态枚举值的映射关系\n- 过滤条件需排除测试账号\n- 金额字段保留两位小数",
        )

    # ----- Tab 4: 开发代码 -----
    with _tab_code:
        st.caption("粘贴参考开发代码（SQL、PySpark、DataWorks 调度等），帮助 AI 理解加工逻辑。")
        dev_code = st.text_area(
            "参考开发代码",
            key="dev_code",
            height=300,
            placeholder="可以粘贴 SQL、PySpark、DataWorks 调度代码等。",
        )

    # =========================
    # 合并 PRD 文本
    # =========================

    prd_text = ""

    if st.session_state.get("uploaded_prd_text", ""):
        prd_text += st.session_state["uploaded_prd_text"]

    if prd_manual_text.strip():
        prd_text += "\n\n【用户手动补充 PRD 内容】\n" + prd_manual_text

    st.session_state["prd_text"] = prd_text

    # =========================
    # 底部统一主按钮 + 材料摘要徽章
    # =========================

    st.divider()

    # 材料摘要徽章
    _badge_parts = []
    if prd_text.strip():
        _badge_parts.append("PRD ✓")
    if meeting_notes.strip():
        _badge_parts.append("会议纪要 ✓")
    if result_table_schema.strip():
        _badge_parts.append("结果表 ✓")
    if source_table_schema.strip():
        _badge_parts.append("源表 ✓")
    if dev_code.strip():
        _badge_parts.append("开发代码 ✓")

    if _badge_parts:
        _badge_html = "  ".join(
            f'<span style="display:inline-block;background:#e8f5e9;color:#2e7d32;'
            f'font-size:0.8em;padding:2px 10px;border-radius:12px;margin:2px;">{b}</span>'
            for b in _badge_parts
        )
        st.markdown(f"已准备材料：{_badge_html}", unsafe_allow_html=True)
    else:
        st.info("尚未填写任何材料，请至少上传或粘贴 PRD 内容。")

    _generate_draft_clicked = st.button(
        "🚀 开始分析 PRD",
        type="primary",
        use_container_width=True
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
                    render_error_with_fold(draft_result)
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
        file_name=_ts_filename("prd_当前需求分析和待确认点", "xlsx"),
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
        st.success("当前没有待确认点，初版分析即为最终版，可直接生成测试用例。")
        st.session_state["prd_pending_answers"] = "无待确认点。"

        if st.button(
            "🚀 直接生成测试用例",
            type="primary"
        ):
            # 无待确认点且无收敛历史 → 初版即终版，跳过第 3 步
            st.session_state["prd_final_analysis_result"] = st.session_state["prd_current_analysis_result"]
            st.session_state["test_case_result"] = ""
            go_to_step(STEP_TEST_CASE)

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
                    render_error_with_fold(new_analysis_result)
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
                render_error_with_fold(final_result)
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
            file_name=_ts_filename("prd_最终版需求提炼表", "xlsx"),
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
                render_error_with_fold(test_result)
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

                # ===== Fragment：SQL 执行区局部刷新，不重渲染整个页面 =====
                @st.fragment(run_every=None)
                def _sql_execution_fragment():
                    _odps_entry_frag = get_odps_entry(
                        st.session_state.get("odps_ak", "").strip(),
                        st.session_state.get("odps_sk", "").strip(),
                        st.session_state.get("odps_project", "").strip(),
                        st.session_state.get("odps_endpoint", "").strip(),
                    )

                    # 一键执行全部 + 停止 + 一键分析
                    _col_batch, _col_stop, _col_analyze = st.columns([2, 1, 2])
                    with _col_batch:
                        _batch_clicked = st.button("一键执行全部 SQL", type="primary", use_container_width=True, key="batch_run_sql")
                    with _col_stop:
                        _stop_clicked = st.button("⏹️ 停止执行", use_container_width=True, key="stop_run_sql",
 help="点击后停止后续 SQL 执行")
                    with _col_analyze:
                        _has_diff = any(
                            v.get("df") is not None and not v.get("df").empty and not v.get("err")
                            for v in st.session_state.get("sql_run_results", {}).values()
                        )
                        _analyze_clicked = st.button("🤖 一键分析所有差异", type="primary", use_container_width=True, key="batch_analyze_diff",
 disabled=not _has_diff)

                    # 停止按钮：清除执行中标志
                    if _stop_clicked:
                        st.session_state["_sql_batch_running"] = False
                        st.rerun()

                    # 一键执行逻辑：用 rerun 驱动的逐段执行，支持中途停止
                    if _batch_clicked:
                        # 启动批量执行模式
                        st.session_state["_sql_batch_running"] = True
                        st.session_state["_sql_batch_idx"] = 0
                        st.rerun()

                    # 批量执行中：每次 rerun 执行一段，检查停止标志
                    if st.session_state.get("_sql_batch_running", False):
                        _idx = st.session_state.get("_sql_batch_idx", 0)
                        if _idx < len(_sql_blocks):
                            _sql_block_b = _sql_blocks[_idx]
                            # 用编辑后的 SQL（如果有）
                            _edit_key_b = f"sql_edit_{_idx + 1}"
                            _sql_to_run = st.session_state.get(_edit_key_b, _sql_block_b)
                            with st.spinner(f"执行 SQL-{_idx + 1:03d}（{_idx + 1}/{len(_sql_blocks)}）..."):
                                _df_r, _err_r = run_single_sql(_odps_entry_frag, _sql_to_run)
                            st.session_state["sql_run_results"][_idx + 1] = {"df": _df_r, "err": _err_r}
                            st.session_state["_sql_batch_idx"] = _idx + 1
                            st.rerun()
                        else:
                            # 全部执行完毕
                            st.session_state["_sql_batch_running"] = False
                            _ok = sum(1 for v in st.session_state["sql_run_results"].values() if not v["err"])
                            _fail = len(st.session_state["sql_run_results"]) - _ok
                            if _fail == 0:
                                st.success(f"全部执行完成（{_ok} 段）。")
                            else:
                                st.warning(f"执行完成：成功 {_ok} 段，失败 {_fail} 段。")
                            st.rerun()

                    # 批量执行中被停止的提示
                    if not st.session_state.get("_sql_batch_running", False) and st.session_state.get("_sql_batch_idx", 0) > 0:
                        _b_idx = st.session_state.get("_sql_batch_idx", 0)
                        if _b_idx < len(_sql_blocks):
                            _ok_s = sum(1 for v in st.session_state.get("sql_run_results", {}).values() if not v["err"])
                            _fail_s = len(st.session_state.get("sql_run_results", {})) - _ok_s
                            st.info(f"执行已停止。已完成 {_b_idx}/{len(_sql_blocks)} 段（成功 {_ok_s}，失败 {_fail_s}）。")
                            st.session_state["_sql_batch_idx"] = 0

                    # 一键分析逻辑
                    if _analyze_clicked:
                        with st.spinner("AI 汇总分析中..."):
                            _batch_content_parts = []
                            for _bi, _bv in st.session_state["sql_run_results"].items():
                                if _bv.get("df") is not None and not _bv.get("df").empty and not _bv.get("err"):
                                    _batch_sql = st.session_state.get(f"sql_edit_{_bi}", _sql_blocks[_bi - 1])
                                    _diff_sample = _bv["df"].head(30).to_csv(index=False)
                                    _batch_content_parts.append(f"""
--- SQL-{_bi:03d} ---

校验 SQL：
```sql
{_batch_sql}
```

差异结果（前 30 行）：
{_diff_sample}
""")
                            _batch_content = "\n".join(_batch_content_parts)
                            _batch_content = f"""
以下是多段校验 SQL 的执行差异结果汇总，共 {len(_batch_content_parts)} 段有差异：

{_batch_content}

请对每段 SQL 的差异原因进行简要分析，并在最后给出整体汇总建议。
"""
                            _batch_analysis = call_llm(SQL_DIFF_ANALYSIS_PROMPT, _batch_content)
                            if is_llm_error(_batch_analysis):
                                render_error_with_fold(_batch_analysis)
                            else:
                                st.session_state["batch_diff_analysis"] = _batch_analysis
                                st.rerun()

                    # 展示汇总分析结果 + 导出
                    _batch_analysis_cached = st.session_state.get("batch_diff_analysis", "")
                    if _batch_analysis_cached:
                        st.divider()
                        _ba_c1, _ba_c2 = st.columns([4, 1])
                        with _ba_c1:
                            st.markdown("#### 🤖 汇总差异分析报告")
                        with _ba_c2:
                            st.download_button(
                                label="📥 导出报告",
                                data=_batch_analysis_cached.encode("utf-8-sig"),
                                file_name=_ts_filename("汇总差异分析报告", "md"),
                                mime="text/markdown",
                                key="download_batch_analysis"
                            )
                        render_markdown_in_scroll_box(
                            title="汇总差异分析",
                            markdown_text=_batch_analysis_cached,
                            height=500,
                            expanded=True
                        )

                    # ===== 结果总览 + 过滤器 =====
                    _run_results = st.session_state.get("sql_run_results", {})
                    _total_sql = len(_sql_blocks)
                    _executed = len(_run_results)
                    _pass_cnt = sum(1 for v in _run_results.values() if not v["err"] and v.get("df") is not None and v["df"].empty)
                    _diff_cnt = sum(1 for v in _run_results.values() if not v["err"] and v.get("df") is not None and not v["df"].empty)
                    _fail_cnt = sum(1 for v in _run_results.values() if v["err"])

                    if _executed > 0:
                        st.markdown(
                            f"**结果总览**：共 {_total_sql} 段 SQL ｜ 已执行 {_executed} ｜ "
                            f"✅ 通过 {_pass_cnt} ｜ ⚠️ 有差异 {_diff_cnt} ｜ ❌ 失败 {_fail_cnt}"
                        )
                        _filter = st.radio(
                            "筛选展示",
                            ["全部", "仅通过", "仅有差异", "仅失败"],
                            horizontal=True,
                            key="sql_result_filter"
                        )
                    else:
                        _filter = "全部"

                    # 逐段展示 + 执行
                    for _i, _sql_block in enumerate(_sql_blocks, 1):
                        # 计算该段状态
                        _cached_f = _run_results.get(_i)
                        if _cached_f:
                            if _cached_f["err"]:
                                _status = "fail"
                            elif _cached_f["df"].empty:
                                _status = "pass"
                            else:
                                _status = "diff"
                        else:
                            _status = "pending"

                        # 过滤判断
                        if _filter == "仅通过" and _status != "pass":
                            continue
                        if _filter == "仅有差异" and _status != "diff":
                            continue
                        if _filter == "仅失败" and _status != "fail":
                            continue

                        _status_icon = {"pass": "✅", "diff": "⚠️", "fail": "❌", "pending": "⏳"}.get(_status, "")

                        with st.expander(f"{_status_icon} SQL-{_i:03d}", expanded=False):
                            # #11 SQL 语法高亮：只读展示 + 编辑切换
                            _edit_key = f"sql_edit_{_i}"
                            if _edit_key not in st.session_state:
                                st.session_state[_edit_key] = _sql_block
                            _edit_mode_key = f"sql_edit_mode_{_i}"
                            if _edit_mode_key not in st.session_state:
                                st.session_state[_edit_mode_key] = False

                            if st.session_state[_edit_mode_key]:
                                # 编辑模式：text_area
                                _edited_sql = st.text_area(
                                    "SQL（编辑中）",
                                    key=_edit_key,
                                    height=300,
                                    label_visibility="collapsed"
                                )
                            else:
                                # 只读模式：带语法高亮的 code block
                                st.code(st.session_state[_edit_key], language="sql", height=300)
                                _edited_sql = st.session_state[_edit_key]

                            _c_run, _c_toggle, _c_reset, _c_export = st.columns([1.5, 1, 1, 1.5])
                            with _c_run:
                                if st.button("▶ 执行", key=f"run_sql_{_i}"):
                                    with st.spinner("执行中..."):
                                        _df_result, _err = run_single_sql(_odps_entry_frag, _edited_sql)
                                    st.session_state["sql_run_results"][_i] = {"df": _df_result, "err": _err}
                                    st.rerun()
                            with _c_toggle:
                                _toggle_label = "✏️ 编辑" if not st.session_state[_edit_mode_key] else "👁️ 只读"
                                if st.button(_toggle_label, key=f"toggle_sql_{_i}"):
                                    st.session_state[_edit_mode_key] = not st.session_state[_edit_mode_key]
                                    st.rerun()
                            with _c_reset:
                                if st.button("🔄 恢复", key=f"reset_sql_{_i}"):
                                    st.session_state[_edit_key] = _sql_block
                                    st.rerun()

                            # 展示已有结果
                            _cached = st.session_state["sql_run_results"].get(_i)
                            if _cached:
                                if _cached["err"]:
                                    render_error_with_fold(f"执行失败：{_cached['err']}")
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
                                            label="📥 CSV",
                                            data=_csv_data,
                                            file_name=_ts_filename(f"sql_{_i:03d}_result", "csv"),
                                            mime="text/csv",
                                            use_container_width=True,
                                            key=f"download_csv_{_i}"
                                        )

                                    # ===== 逐段 AI 差异分析 =====
                                    _analysis_key = f"sql_diff_analysis_{_i}"
                                    _a_c1, _a_c2 = st.columns([3, 1])
                                    with _a_c1:
                                        if st.button("🤖 AI 分析差异原因", key=f"analyze_diff_{_i}"):
                                            with st.spinner("AI 分析中..."):
                                                _diff_sample = _cached["df"].head(50).to_csv(index=False)
                                                _analysis_content = f"""
以下是执行的校验 SQL：

```sql
{_edited_sql}
```

以下是 SQL 执行返回的差异结果（前 50 行）：

{_diff_sample}

请分析差异原因并给出建议。分析要点：
1. 差异数据的整体特征（差异行数、涉及哪些字段、差异值的模式）
2. 可能的差异原因（如：口径不一致、分区取值不同、关联条件导致数据放大/丢失、NULL 处理差异、枚举映射错误、金额精度问题等）
3. 需要进一步确认的问题
4. 修复建议
"""
                                                _analysis_result = call_llm(
                                                    SQL_DIFF_ANALYSIS_PROMPT,
                                                    _analysis_content
                                                )
                                                if is_llm_error(_analysis_result):
                                                    render_error_with_fold(_analysis_result)
                                                else:
                                                    st.session_state[_analysis_key] = _analysis_result
                                                    st.rerun()

                                    _analysis_cached = st.session_state.get(_analysis_key, "")
                                    if _analysis_cached:
                                        with _a_c2:
                                            st.download_button(
                                                label="📥 导出",
                                                data=_analysis_cached.encode("utf-8-sig"),
                                                file_name=_ts_filename(f"sql_{_i:03d}_差异分析报告", "md"),
                                                mime="text/markdown",
                                                key=f"download_analysis_{_i}"
                                            )
                                        st.divider()
                                        st.markdown("#### 🤖 AI 差异分析报告")
                                        render_markdown_in_scroll_box(
                                            title="差异分析",
                                            markdown_text=_analysis_cached,
                                            height=400,
                                            expanded=True
                                        )

                    # 清空执行结果 — 放在折叠区，防止误触
                    with st.expander("⚠️ 危险操作（清空所有执行结果和分析报告）", expanded=False):
                        st.warning("点击后会清空所有 SQL 的执行结果和 AI 差异分析报告，不可恢复。")
                        if st.button("确认清空全部执行结果", key="clear_run_results"):
                            st.session_state["sql_run_results"] = {}
                            st.session_state["batch_diff_analysis"] = ""
                            st.session_state["_sql_batch_running"] = False
                            st.session_state["_sql_batch_idx"] = 0
                            for _ci in range(1, len(_sql_blocks) + 1):
                                st.session_state.pop(f"sql_diff_analysis_{_ci}", None)
                            st.rerun()

                _sql_execution_fragment()

        else:
            st.info("未配置 ODPS 连接，仅支持下载 SQL 脚本手动执行。配置方法见侧边栏 ODPS 连接配置。")
else:
    st.warning("当前步骤状态异常，已返回第 1 步。")
    st.session_state["current_step"] = STEP_INPUT
    st.rerun()
