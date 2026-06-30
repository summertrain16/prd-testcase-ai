"""PyODPS 连接器工具函数。

封装 MaxCompute/DataWorks 的连接、表结构拉取、数据预览、SQL 执行能力。
被 app.py 调用，用于替代手动上传 xlsx 表结构，并支持在线执行生成的校验 SQL。
"""

import streamlit as st
import pandas as pd
from odps import ODPS

MAX_RESULT_ROWS = 1000


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
    return ODPS(ak, sk, project=project, endpoint=endpoint)


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
    sql_clean = sql_text.strip()
    while sql_clean.endswith(";"):
        sql_clean = sql_clean[:-1].rstrip()
    # 去除 SQL 前后的注释外多余空行
    sql_clean = sql_clean.strip()

    if not sql_clean:
        return pd.DataFrame(), "SQL 内容为空"

    try:
        instance = odps_entry.execute_sql(sql_clean)
    except Exception as e:
        return pd.DataFrame(), f"SQL 执行失败：{e}"

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
