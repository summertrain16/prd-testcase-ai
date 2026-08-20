"""PyODPS 连接器工具函数。

封装 MaxCompute/DataWorks 的连接、表结构拉取、数据预览、SQL 执行能力。
被 app.py 调用，用于替代手动上传 xlsx 表结构，并支持在线执行生成的校验 SQL。
"""

import json
import streamlit as st
import pandas as pd
from odps import ODPS

MAX_RESULT_ROWS = 1000

SESSION_TABLE = "ods_shinebed_dev.ods_datatest_session_history"


@st.cache_resource
def get_odps_entry(ak, sk, project, endpoint):
    """创建 ODPS 连接对象（带缓存，同参数只建一次）。

    参数：
        ak: AccessKey ID
        sk: AccessKey Secret
        project: 默认 MaxCompute 项目名（可留空，表名带项目前缀即可跨项目）
        endpoint: MaxCompute endpoint，例如 http://service.odps.aliyun.com/api

    返回：
        ODPS 连接对象；ak/sk/endpoint 三项任一为空则返回 None。project 可空。
    """
    if not all([ak, sk, endpoint]):
        return None
    project = project.strip() if project else None
    o = ODPS(ak, sk, project=project, endpoint=endpoint)
    # 把原始凭证存到自定义属性上，run_single_sql 重建连接时用
    o._user_ak = ak
    o._user_sk = sk
    o._user_project = project
    o._user_endpoint = endpoint
    return o


def test_odps_connection(odps_entry):
    """测试 ODPS 连接是否可用，返回 (是否成功, 错误信息)。"""
    if odps_entry is None:
        return False, "连接对象为空，请检查配置"
    try:
        odps_entry.projects
        return True, "连接成功"
    except Exception as e:
        return False, str(e)


def get_table_schema_text(odps_entry, table_name):
    """拉取表结构，返回和 xlsx 解析后格式一致的 Markdown 文本。

    输出格式和 render_table_schema_uploader 中的 schema_text 兼容，
    可直接拼入 source_table_schema / result_table_schema，下游 LLM 调用无需改动。

    参数：
        odps_entry: ODPS 连接对象
        table_name: 表名

    返回：
        Markdown 格式的表结构文本。
    """
    t = odps_entry.get_table(table_name)
    parts = []
    parts.append(f"### 表结构（来自 MaxCompute）：{table_name}")
    if t.comment:
        parts.append(f"表注释：{t.comment}")

    parts.append("")
    parts.append("| 字段名 | 类型 | 注释 |")
    parts.append("|---|---|---|")

    # 普通列
    for col in t.schema.columns:
        parts.append(f"| {col.name} | {col.type} | {col.comment or ''} |")

    # 分区列
    if t.schema.partitions:
        parts.append("")
        parts.append("分区字段：")
        for col in t.schema.partitions:
            parts.append(f"| {col.name} (分区) | {col.type} | {col.comment or ''} |")
    else:
        parts.append("")
        parts.append("分区：无")

    return "\n".join(parts)


def preview_table_data(odps_entry, table_name, partition="", limit=20):
    """预览表数据，返回 DataFrame。

    用 tunnel reader 读取，绕过 execute_sql open_reader 的 schema 问题。

    参数：
        odps_entry: ODPS 连接对象
        table_name: 表名
        partition: 分区条件，例如 "pt='20250101'"。留空则不加 WHERE。
        limit: 返回行数，默认 20

    返回：
        pandas DataFrame。
    """
    t = odps_entry.get_table(table_name)

    # 构造分区过滤条件
    partition_spec = partition.strip() if partition.strip() else None

    with t.open_reader(partition=partition_spec) as reader:
        # 从表 schema 取列名（包含普通列，不含分区列）
        col_names = [col.name for col in t.schema.columns]

        # 分区列也要加上（tunnel reader 返回的行里包含分区列值）
        if t.schema.partitions:
            for pcol in t.schema.partitions:
                if pcol.name not in col_names:
                    col_names.append(pcol.name)

        rows = []
        for i, record in enumerate(reader):
            if i >= limit:
                break
            row = []
            for col_name in col_names:
                try:
                    val = record[col_name]
                except Exception:
                    val = None
                row.append(val if val is not None else "")
            rows.append(row)

    return pd.DataFrame(rows, columns=col_names)


def is_partitioned_table(odps_entry, table_name):
    """判断表是否有分区字段，返回 bool。

    需要 Describe 权限。如果权限不足或出错，返回 None 表示无法判断。
    """
    try:
        t = odps_entry.get_table(table_name)
        return bool(t.schema.partitions)
    except Exception:
        return None


def run_single_sql(odps_entry, sql_text, max_rows=MAX_RESULT_ROWS):
    """执行单段 SQL，返回 (DataFrame, 错误信息)。

    多路兜底读取结果：
    1. instance.to_pandas() — PyODPS 0.12.0+ 内置方法，自动处理 schema
    2. open_reader(tunnel=True).to_pandas() — Instance Tunnel 通道
    3. open_reader(tunnel=False) — 旧 Results 接口

    结果超过 max_rows 行会自动截断，避免撑爆内存。

    参数：
        odps_entry: ODPS 连接对象
        sql_text: 单段 SQL 文本
        max_rows: 最大返回行数，默认 1000

    返回：
        (result_df, error_msg)
        - 成功且无数据：返回 (空 DataFrame, None)
        - 成功且有数据：返回 (DataFrame, None)，超过 max_rows 会截断
        - 失败：返回 (空 DataFrame, 错误字符串)
    """
    # 预处理 SQL 文本：
    # 1. 去除尾部分号（MaxCompute execute_sql 不接受以 ; 结尾，会把分号后空内容当第二语句报错）
    # 2. 去除多余空白行
    # 3. 如果 SQL 以块注释 /* ... */ 开头，execute_sql 可能不认，需要清理掉纯注释行
    sql_clean = sql_text.strip()
    while sql_clean.endswith(";"):
        sql_clean = sql_clean[:-1].rstrip()

    # 如果整段 SQL 被包在 /* ... */ 注释里（某些 LLM 输出会这样），提取注释后的 SQL
    # 但保留行内注释，只处理"整段以 /* 开头且前面没有 SQL 语句"的情况
    lines = sql_clean.split("\n")
    # 检查是否开头是一整块注释（连续 /* 到 */ ），后面才是 SQL
    stripped_first = lines[0].strip() if lines else ""
    if stripped_first.startswith("/*"):
        # 找 */ 结束行
        comment_end_idx = None
        for idx, line in enumerate(lines):
            if "*/" in line:
                comment_end_idx = idx
                break
        if comment_end_idx is not None and comment_end_idx < len(lines) - 1:
            # 注释块后面还有内容，去掉注释块
            sql_clean = "\n".join(lines[comment_end_idx + 1:]).strip()

    # 去除行首行尾空白
    sql_clean = sql_clean.strip()

    if not sql_clean:
        return pd.DataFrame(), "SQL 内容为空"

    # 确定用哪个 project 提交 SQL 任务
    # 优先级：用户在侧边栏配置的 project > 从 SQL 提取的项目名
    # 原因：execute_sql 提交任务需要对项目有 CreateInstance 权限，
    # 用户配置的 project 通常是有执行权限的项目，
    # 而 SQL 中 FROM 的第一个项目可能只是源表所在项目，不一定有执行权限
    import re as _re
    _sql_projects = set(_re.findall(r'\b(?:FROM|JOIN)\s+(\w+)\.\w+', sql_clean, _re.IGNORECASE))

    _exec_odps = odps_entry
    _user_project = getattr(odps_entry, '_user_project', None)
    _user_ak = getattr(odps_entry, '_user_ak', None)
    _user_sk = getattr(odps_entry, '_user_sk', None)
    _user_endpoint = getattr(odps_entry, '_user_endpoint', None)

    # 只有用户没配 project 时，才从 SQL 提取（兜底）
    if not _user_project and _sql_projects:
        _sql_project = next(iter(_sql_projects))  # 取第一个
        _exec_odps = ODPS(
            _user_ak,
            _user_sk,
            project=_sql_project,
            endpoint=_user_endpoint
        )
        _exec_odps._user_ak = _user_ak
        _exec_odps._user_sk = _user_sk
        _exec_odps._user_project = _sql_project
        _exec_odps._user_endpoint = _user_endpoint
    # 用户配了 project 时，直接用，不管 SQL 里查的是哪个项目

    try:
        instance = _exec_odps.execute_sql(sql_clean)
    except Exception as e:
        _err_msg = str(e)
        # 如果 execute_sql 报错，尝试用 run_sql + wait_for_success（异步方式有时能绕过）
        try:
            instance = _exec_odps.run_sql(sql_clean)
            instance.wait_for_success()
        except Exception as e2:
            _sql_proj_display = ', '.join(_sql_projects) if _sql_projects else '(无)'
            _proj_info = f"配置的 project: {_user_project or '(空)'}, SQL 涉及的项目: {_sql_proj_display}"
            return pd.DataFrame(), (
                f"SQL 执行失败（同步）：{_err_msg}\n"
                f"SQL 执行失败（异步）：{e2}\n"
                f"{_proj_info}\n"
                f"提示：如果报权限错（NoPermission / CreateInstance），请在侧边栏 ODPS 配置的 Project 栏填写你有执行权限的项目名"
                f"--- 完整 SQL 文本 ---\n{sql_clean}"
            )

    # 方案 1：instance.to_pandas()（PyODPS 0.12.0+，内部自动处理 schema）
    try:
        df = instance.to_pandas()
        if df is None or df.empty:
            return pd.DataFrame(), None
        truncated = len(df) > max_rows
        if truncated:
            df = df.head(max_rows)
            st.warning(f"结果共 {len(df)} 行，已截断为前 {max_rows} 行。如需完整结果请导出 CSV 或优化 SQL。")
        return df, None
    except Exception as e1:
        _err1 = str(e1)

    # 方案 2：open_reader(tunnel=True) + reader.to_pandas()
    try:
        with instance.open_reader(tunnel=True) as reader:
            if reader.count == 0:
                return pd.DataFrame(), None
            df = reader.to_pandas()
            if df is None or df.empty:
                return pd.DataFrame(), None
            truncated = len(df) > max_rows
            if truncated:
                df = df.head(max_rows)
                st.warning(f"结果共 {len(df)} 行，已截断为前 {max_rows} 行。如需完整结果请导出 CSV 或优化 SQL。")
            return df, None
    except Exception as e2:
        _err2 = str(e2)

    # 方案 3：open_reader(tunnel=False) 旧 Results 接口 + 手动读取
    try:
        with instance.open_reader(tunnel=False) as reader:
            total = reader.count
            if total == 0:
                return pd.DataFrame(), None

            # 旧接口取列名，包 try-except
            try:
                cols = [c.name for c in reader._schema.columns]
            except Exception:
                cols = None

            rows = []
            first_record = None
            for i, record in enumerate(reader):
                if i >= max_rows:
                    break
                if i == 0:
                    first_record = record
                # 不用 record.values（可能触发 None），改用按列名逐个取值
                if cols:
                    row = []
                    for col_name in cols:
                        try:
                            row.append(record[col_name])
                        except Exception:
                            row.append(None)
                    rows.append(row)
                else:
                    # 没有列名时才用 values
                    rows.append(record.values)

            # 兜底取列名
            if not cols and first_record is not None:
                try:
                    cols = list(first_record.keys())
                except Exception:
                    cols = [f"col_{j+1}" for j in range(len(rows[0]) if rows else 0)]

            truncated = total > max_rows
            df = pd.DataFrame(rows, columns=cols)
            if truncated:
                st.warning(f"结果共 {total} 行，已截断为前 {max_rows} 行。如需完整结果请导出 CSV 或优化 SQL。")
            return df, None
    except Exception as e3:
        _err3 = str(e3)

    # 全部失败
    return pd.DataFrame(), f"SQL 执行成功但结果读取失败（3 种方式均失败）：\n1) to_pandas: {_err1}\n2) tunnel reader: {_err2}\n3) legacy reader: {_err3}"


# =========================
# 会话历史记录持久化
# =========================

def _sanitize_prd_name(name: str) -> str:
    """清洗 PRD 名称：限 50 字符，过滤危险特殊字符。"""
    if not name:
        return ""
    _dangerous = set('"\'`;\\')
    _cleaned = "".join(c for c in name if c not in _dangerous)
    return _cleaned.strip()[:50]


def _escape_sql_string(value: str) -> str:
    """转义 SQL 字符串中的单引号，防止注入。"""
    if value is None:
        return ""
    return str(value).replace("'", "''")


def save_session_to_odps(odps_entry, session_data: dict, is_new_version: bool = False):
    """保存会话到 ODPS 表（用 INSERT OVERWRITE 替代 UPDATE/DELETE）。

    参数：
        odps_entry: ODPS 连接对象
        session_data: 包含以下 key 的 dict
            - session_id, prd_name, current_step
            - prd_text, meeting_notes, dev_code
            - result_schema, source_schema (JSON 字符串)
            - draft_analysis, pending_points (JSON 字符串)
            - pending_answers, pending_history, ignored_pending_points
            - final_analysis, test_cases
            - sql_results (JSON 字符串)
        is_new_version: True=新建版本(旧记录is_latest置false+新记录)，
                        False=更新当前最新版本(删旧最新+插新)

    返回：
        (是否成功, 错误信息)
    """
    if odps_entry is None:
        return False, "ODPS 连接未配置"

    _sid = _escape_sql_string(session_data.get("session_id", ""))
    if not _sid:
        return False, "session_id 为空"

    _now = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")

    # 所有字段名
    _cols = [
        "session_id", "prd_name", "version", "is_latest", "current_step",
        "prd_text", "meeting_notes", "result_schema", "source_schema", "dev_code",
        "draft_analysis", "pending_points", "pending_answers", "pending_history",
        "ignored_pending_points", "final_analysis", "test_cases", "sql_results",
        "create_time", "update_time"
    ]

    try:
        if is_new_version:
            # 查当前最大版本号
            _sql_max_ver = (
                f"SELECT MAX(version) AS max_v FROM {SESSION_TABLE} "
                f"WHERE session_id = '{_sid}'"
            )
            _instance = odps_entry.execute_sql(_sql_max_ver)
            _max_ver = 0
            with _instance.open_reader() as _r:
                for _rec in _r:
                    _max_ver = _rec[0] if _rec[0] else 0
                    break
            _new_ver = int(_max_ver) + 1
            _create_time = _now

            # INSERT OVERWRITE：旧记录 is_latest 全置 false（匹配 session_id 的），
            # 其他记录原样保留，最后 UNION ALL 新记录
            # 注意：列顺序必须和表结构一致：session_id, prd_name, version, is_latest, ...
            _case_cols = ", ".join([
                "session_id", "prd_name",
                "version",
                "CASE WHEN session_id = '{}' THEN false ELSE is_latest END AS is_latest".format(_sid),
                "current_step", "prd_text", "meeting_notes",
                "result_schema", "source_schema", "dev_code",
                "draft_analysis", "pending_points", "pending_answers",
                "pending_history", "ignored_pending_points",
                "final_analysis", "test_cases", "sql_results",
                "create_time", "update_time"
            ])

            _vals = _build_values_clause(session_data, _new_ver, True, _create_time, _now)

            _sql_overwrite = (
                f"INSERT OVERWRITE TABLE {SESSION_TABLE} "
                f"SELECT {_case_cols} FROM {SESSION_TABLE} "
                f"UNION ALL "
                f"SELECT {_vals}"
            )
            odps_entry.execute_sql(_sql_overwrite)

            return True, f"保存成功（版本 {_new_ver}）"

        else:
            # 更新当前最新版本：删旧最新记录 + 插新
            # 查该 session 的 version 号
            _sql_ver = (
                f"SELECT MAX(version) AS max_v FROM {SESSION_TABLE} "
                f"WHERE session_id = '{_sid}' AND is_latest = true"
            )
            _instance = odps_entry.execute_sql(_sql_ver)
            _new_ver = 1
            with _instance.open_reader() as _r:
                for _rec in _r:
                    _new_ver = int(_rec[0]) if _rec[0] else 1
                    break
            _create_time = session_data.get("create_time", _now)

            _vals = _build_values_clause(session_data, _new_ver, True, _create_time, _now)

            # INSERT OVERWRITE：保留除当前 is_latest 外的所有记录 + 新记录
            _all_cols = ", ".join(_cols)
            _sql_overwrite = (
                f"INSERT OVERWRITE TABLE {SESSION_TABLE} "
                f"SELECT {_all_cols} FROM {SESSION_TABLE} "
                f"WHERE NOT (session_id = '{_sid}' AND is_latest = true) "
                f"UNION ALL "
                f"SELECT {_vals}"
            )
            odps_entry.execute_sql(_sql_overwrite)

            return True, f"保存成功（版本 {_new_ver}）"

    except Exception as e:
        return False, f"保存失败：{e}"


def _build_values_clause(session_data: dict, version: int, is_latest: bool, create_time: str, update_time: str) -> str:
    """构造 INSERT 用的 VALUES 子句（SELECT ... 形式，每列用引号包裹）。"""
    _vals = [
        _escape_sql_string(session_data.get("session_id", "")),
        _escape_sql_string(session_data.get("prd_name", "")),
        str(version),
        "true" if is_latest else "false",
        _escape_sql_string(session_data.get("current_step", "")),
        _escape_sql_string(session_data.get("prd_text", "")),
        _escape_sql_string(session_data.get("meeting_notes", "")),
        _escape_sql_string(session_data.get("result_schema", "")),
        _escape_sql_string(session_data.get("source_schema", "")),
        _escape_sql_string(session_data.get("dev_code", "")),
        _escape_sql_string(session_data.get("draft_analysis", "")),
        _escape_sql_string(session_data.get("pending_points", "")),
        _escape_sql_string(session_data.get("pending_answers", "")),
        _escape_sql_string(session_data.get("pending_history", "")),
        _escape_sql_string(session_data.get("ignored_pending_points", "")),
        _escape_sql_string(session_data.get("final_analysis", "")),
        _escape_sql_string(session_data.get("test_cases", "")),
        _escape_sql_string(session_data.get("sql_results", "")),
        _escape_sql_string(create_time),
        _escape_sql_string(update_time),
    ]
    # 布尔和数字不加引号
    _quoted = []
    for i, v in enumerate(_vals):
        if i in (2,):  # version 是数字
            _quoted.append(v)
        elif i in (3,):  # is_latest 是布尔
            _quoted.append(v)
        else:
            _quoted.append(f"'{v}'")
    return ", ".join(_quoted)


def load_session_list(odps_entry, limit=50):
    """加载历史记录列表（只返回 is_latest=true 的摘要信息）。

    返回：
        [{"session_id", "prd_name", "version", "update_time"}, ...]
        失败返回空列表。
    """
    if odps_entry is None:
        return []

    try:
        _sql = (
            f"SELECT session_id, prd_name, version, update_time "
            f"FROM {SESSION_TABLE} "
            f"WHERE is_latest = true "
            f"ORDER BY update_time DESC "
            f"LIMIT {limit}"
        )
        _instance = odps_entry.execute_sql(_sql)
        _rows = []
        with _instance.open_reader() as _r:
            for _rec in _r:
                _rows.append({
                    "session_id": _rec[0],
                    "prd_name": _rec[1],
                    "version": int(_rec[2]) if _rec[2] else 1,
                    "update_time": _rec[3],
                })
        return _rows
    except Exception:
        return []


def load_session_detail(odps_entry, session_id, version=None):
    """加载某条会话的完整详情。

    参数：
        version: 不传则加载 is_latest=true 的；传了加载指定版本

    返回：
        dict 包含全部字段，失败返回 None。
    """
    if odps_entry is None:
        return None

    _sid = _escape_sql_string(session_id)
    if not _sid:
        return None

    try:
        if version:
            _sql = (
                f"SELECT * FROM {SESSION_TABLE} "
                f"WHERE session_id = '{_sid}' AND version = {int(version)}"
            )
        else:
            _sql = (
                f"SELECT * FROM {SESSION_TABLE} "
                f"WHERE session_id = '{_sid}' AND is_latest = true"
            )

        _instance = odps_entry.execute_sql(_sql)
        with _instance.open_reader() as _r:
            for _rec in _r:
                return {
                    "session_id": _rec[0] or "",
                    "prd_name": _rec[1] or "",
                    "version": int(_rec[2]) if _rec[2] else 1,
                    "is_latest": bool(_rec[3]),
                    "current_step": _rec[4] or "",
                    "prd_text": _rec[5] or "",
                    "meeting_notes": _rec[6] or "",
                    "result_schema": _rec[7] or "",
                    "source_schema": _rec[8] or "",
                    "dev_code": _rec[9] or "",
                    "draft_analysis": _rec[10] or "",
                    "pending_points": _rec[11] or "",
                    "pending_answers": _rec[12] or "",
                    "pending_history": _rec[13] or "",
                    "ignored_pending_points": _rec[14] or "",
                    "final_analysis": _rec[15] or "",
                    "test_cases": _rec[16] or "",
                    "sql_results": _rec[17] or "",
                    "create_time": _rec[18] or "",
                    "update_time": _rec[19] or "",
                }
        return None
    except Exception:
        return None


def load_session_versions(odps_entry, session_id):
    """查某 PRD 的所有版本列表。

    返回：
        [{"version", "create_time", "is_latest"}, ...]
        失败返回空列表。
    """
    if odps_entry is None:
        return []

    _sid = _escape_sql_string(session_id)
    if not _sid:
        return []

    try:
        _sql = (
            f"SELECT version, create_time, is_latest "
            f"FROM {SESSION_TABLE} "
            f"WHERE session_id = '{_sid}' "
            f"ORDER BY version DESC"
        )
        _instance = odps_entry.execute_sql(_sql)
        _rows = []
        with _instance.open_reader() as _r:
            for _rec in _r:
                _rows.append({
                    "version": int(_rec[0]) if _rec[0] else 1,
                    "create_time": _rec[1] or "",
                    "is_latest": bool(_rec[2]),
                })
        return _rows
    except Exception:
        return []


def delete_session(odps_entry, session_id, version=None):
    """删除会话记录。

    参数：
        version: 不传删全版本，传了删单版本

    返回：
        (是否成功, 错误信息)
    """
    if odps_entry is None:
        return False, "ODPS 连接未配置"

    _sid = _escape_sql_string(session_id)
    if not _sid:
        return False, "session_id 为空"

    try:
        _all_cols = (
            "session_id, prd_name, version, is_latest, current_step, "
            "prd_text, meeting_notes, result_schema, source_schema, dev_code, "
            "draft_analysis, pending_points, pending_answers, pending_history, "
            "ignored_pending_points, final_analysis, test_cases, sql_results, "
            "create_time, update_time"
        )
        if version:
            _sql = (
                f"INSERT OVERWRITE TABLE {SESSION_TABLE} "
                f"SELECT {_all_cols} FROM {SESSION_TABLE} "
                f"WHERE NOT (session_id = '{_sid}' AND version = {int(version)})"
            )
        else:
            _sql = (
                f"INSERT OVERWRITE TABLE {SESSION_TABLE} "
                f"SELECT {_all_cols} FROM {SESSION_TABLE} "
                f"WHERE session_id <> '{_sid}'"
            )
        odps_entry.execute_sql(_sql)
        return True, "删除成功"
    except Exception as e:
        return False, f"删除失败：{e}"
