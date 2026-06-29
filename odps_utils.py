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

    对于分区表，必须传 partition 参数（例如 pt='20250101'），
    否则全表扫描会触发底层 "Unsupported getitem value: None" 错误。

    参数：
        odps_entry: ODPS 连接对象
        table_name: 表名
        partition: 分区条件，例如 "pt='20250101'"。留空则不加 WHERE。
        limit: 返回行数，默认 20

    返回：
        pandas DataFrame。
    """
    if partition.strip():
        sql = f"SELECT * FROM {table_name} WHERE {partition.strip()} LIMIT {limit};"
    else:
        sql = f"SELECT * FROM {table_name} LIMIT {limit};"
    instance = odps_entry.execute_sql(sql)
    with instance.open_reader() as reader:
        rows = [r.values for r in reader]
        cols = [c.name for c in reader._schema.columns]
    return pd.DataFrame(rows, columns=cols)


def is_partitioned_table(odps_entry, table_name):
    """判断表是否有分区字段，返回 bool。"""
    t = odps_entry.get_table(table_name)
    return bool(t.schema.partitions)


def run_single_sql(odps_entry, sql_text, max_rows=MAX_RESULT_ROWS):
    """执行单段 SQL，返回 (DataFrame, 错误信息)。

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
    try:
        instance = odps_entry.execute_sql(sql_text)
        with instance.open_reader() as reader:
            total = reader.count
            if total == 0:
                return pd.DataFrame(), None

            cols = [c.name for c in reader._schema.columns]
            truncated = total > max_rows

            rows = []
            for i, record in enumerate(reader):
                if i >= max_rows:
                    break
                rows.append(record.values)

            df = pd.DataFrame(rows, columns=cols)
            if truncated:
                st.warning(f"结果共 {total} 行，已截断为前 {max_rows} 行。如需完整结果请导出 CSV 或优化 SQL。")
            return df, None
    except Exception as e:
        return pd.DataFrame(), str(e)
