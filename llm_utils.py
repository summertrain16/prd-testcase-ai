"""
LLM 调用工具方法。

包含：
- call_llm: 调用大模型，返回文本结果
- is_llm_error: 判断大模型返回是否为错误信息
"""

import streamlit as st
from openai import OpenAI


def call_llm(system_prompt: str, user_content: str) -> str:
    """
    调用大模型，返回文本结果。
    兼容 OpenAI SDK 标准对象、dict、字符串返回。
    从 session_state 读取 API Key / Base URL / Model，支持每个用户自己配置。
    """
    api_key = st.session_state.get("user_openai_api_key", "").strip()
    base_url = st.session_state.get("user_openai_base_url", "").strip()
    model = st.session_state.get("user_openai_model", "").strip()

    if not api_key:
        return "错误：未配置 API Key，请在左侧侧边栏填写你的 API Key。"

    if not base_url:
        base_url = "https://api.openai.com/v1"
    if not model:
        model = "gpt-4o-mini"

    try:
        llm_client = OpenAI(
            api_key=api_key,
            base_url=base_url
        )

        response = llm_client.chat.completions.create(
            model=model,
            messages=[
                {
                    "role": "system",
                    "content": system_prompt
                },
                {
                    "role": "user",
                    "content": user_content
                }
            ],
            temperature=0.2
        )

        if isinstance(response, str):
            if response.strip().lower().startswith("<!doctype html") or "<html" in response.lower():
                return (
                    "调用大模型失败：接口返回了 HTML 页面，而不是模型结果。\n\n"
                    "这通常说明 Base URL 配置错误。\n"
                    "如果你使用 New API，请把 Base URL 改成：\n\n"
                    "http://你的NewAPI地址/v1\n\n"
                    "注意不要写成后台首页地址，也不要写成 /v1/chat/completions。"
                )
            return response

        if isinstance(response, dict):
            try:
                return response["choices"][0]["message"]["content"]
            except Exception:
                return str(response)

        if hasattr(response, "choices"):
            return response.choices[0].message.content

        return str(response)

    except Exception as e:
        return f"调用大模型失败：{str(e)}"


def is_llm_error(result: str) -> bool:
    """
    判断大模型返回是否为错误信息。
    """
    if not result:
        return True

    text = str(result).strip()

    return (
        text.startswith("错误：")
        or text.startswith("调用大模型失败：")
    )
