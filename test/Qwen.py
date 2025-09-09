#!/usr/bin/env python3
"""
Qwen API 测试示例
=================

本文件演示如何使用阿里云百炼平台的Qwen大模型API进行基本的对话交互。
该示例展示了如何配置API客户端、发送消息以及处理流式响应。

依赖：
    - openai: OpenAI Python SDK，用于兼容阿里云百炼API
    - os: 用于读取环境变量

环境变量：
    DASHSCOPE_API_KEY: 阿里云百炼API密钥

作者：BidAnalysis Team
创建时间：2025年
"""

from openai import OpenAI
import os

# 初始化OpenAI客户端，用于连接阿里云百炼平台
# 使用OpenAI SDK的兼容模式来调用阿里云百炼API
client = OpenAI(
    # 从环境变量中获取API密钥，确保安全性
    # 如果没有配置环境变量，请用阿里云百炼API Key替换：api_key="sk-xxx"
    api_key=os.getenv("DASHSCOPE_API_KEY"),
    # 阿里云百炼平台的兼容模式API基础URL
    base_url="https://dashscope.aliyuncs.com/compatible-mode/v1",
)

# 构建对话消息列表
# messages 格式遵循OpenAI的对话格式标准
messages = [{"role": "user", "content": "你是谁"}]

# 创建对话完成请求
# 使用流式响应来实时获取模型的思考过程和回答
completion = client.chat.completions.create(
    model="qwen-plus-2025-04-28",  # 使用Qwen Plus模型，您可以按需更换为其它深度思考模型
    messages=messages,             # 对话消息列表
    # enable_thinking 参数开启思考过程显示
    # 注意：qwen3-30b-a3b-thinking-2507、qwen3-235b-a22b-thinking-2507、QwQ 与 DeepSeek-R1 模型总会进行思考，不支持该参数
    extra_body={"enable_thinking": True},
    stream=True,                   # 启用流式响应，实时获取生成内容
    # stream_options={
    #     "include_usage": True    # 可选：包含使用量统计信息
    # },
)

# 初始化变量用于存储响应内容
reasoning_content = ""  # 完整思考过程内容
answer_content = ""     # 完整回答内容
is_answering = False    # 标志是否进入回答阶段

# 打印思考过程开始标志
print("\n" + "=" * 20 + "思考过程" + "=" * 20 + "\n")

# 处理流式响应
# 遍历每个响应块，实时显示思考过程和最终回答
for chunk in completion:
    # 检查是否包含选择数据，如果没有则可能是使用量统计信息
    if not chunk.choices:
        print("\nUsage:")  # 打印使用量统计
        print(chunk.usage)
        continue

    # 获取响应增量数据
    delta = chunk.choices[0].delta

    # 处理思考过程内容（reasoning_content）
    # 只在回答阶段开始前收集和显示思考内容
    if hasattr(delta, "reasoning_content") and delta.reasoning_content is not None:
        if not is_answering:
            # 实时打印思考过程，不换行，立即刷新输出
            print(delta.reasoning_content, end="", flush=True)
        # 累积完整的思考过程内容
        reasoning_content += delta.reasoning_content

    # 处理最终回答内容（content）
    # 当收到content时，表示思考结束，开始输出最终回答
    if hasattr(delta, "content") and delta.content:
        if not is_answering:
            # 打印回答阶段开始标志
            print("\n" + "=" * 20 + "完整回复" + "=" * 20 + "\n")
            is_answering = True
        # 实时打印回答内容，不换行，立即刷新输出
        print(delta.content, end="", flush=True)
        # 累积完整的回答内容
        answer_content += delta.content