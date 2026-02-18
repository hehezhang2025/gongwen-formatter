#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
配置文件
"""

# Ollama 配置
OLLAMA_CONFIG = {
    "base_url": "http://localhost:11434",
    "model": "qwen2.5:7b",
    "temperature": 0.1,  # 低温度保证稳定性
    "timeout": 120  # 超时时间（秒）
}

# 处理模式
PROCESSING_MODE = {
    "ORIGINAL": "original",  # 仅原有格式化
    "LLM": "llm",           # 仅LLM增强
    "BOTH": "both"          # 同时生成两个（默认）
}
