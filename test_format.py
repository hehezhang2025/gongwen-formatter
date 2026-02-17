#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""测试格式调整"""

import sys
sys.path.insert(0, '/Users/weipengzhang/Documents/个人/学习笔记/codebuddy/codebuddy_gongwen')

from gongwen_formatter_cli import format_gongwen

# 测试文件
test_file = "测试文档_全错版.docx"
output_file = format_gongwen(test_file)
if output_file:
    print(f"\n✅ 成功生成: {output_file}")
else:
    print("\n❌ 处理失败")
