#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LLM功能测试脚本
用于快速测试Ollama连接和LLM识别效果
"""

import sys
import os

print("\n" + "="*60)
print("  🧪 LLM功能测试脚本")
print("="*60)

# 测试1: 导入模块
print("\n1️⃣ 测试模块导入...")
try:
    from config import OLLAMA_CONFIG
    from llm_client import OllamaClient
    from llm_formatter import llm_format_document
    print("   ✅ 所有模块导入成功")
except ImportError as e:
    print(f"   ❌ 模块导入失败: {e}")
    sys.exit(1)

# 测试2: Ollama连接
print("\n2️⃣ 测试Ollama连接...")
print(f"   配置: {OLLAMA_CONFIG['base_url']}")
print(f"   模型: {OLLAMA_CONFIG['model']}")

client = OllamaClient()
success, message = client.check_connection()
print(f"   {message}")

if not success:
    print("\n❌ Ollama连接失败，请检查:")
    print("   1. Ollama是否已安装: https://ollama.com")
    print("   2. Ollama是否在运行: ollama serve")
    print("   3. 模型是否已下载: ollama pull qwen2.5:7b")
    sys.exit(1)

# 测试3: 简单文档识别
print("\n3️⃣ 测试文档识别能力...")
test_doc = """0: 关于加强项目管理的通知
1: 各部门：
2: 为了提高项目管理水平，现就有关事项通知如下。
3: 一、加强组织领导
4: 各部门要高度重视项目管理工作。
5: （一）成立工作小组
6: 由部门负责人担任组长，组织实施。
7: 1.明确责任分工
8: 每个成员职责明确，分工协作。
9: XX科技有限公司
10: 2025年2月17日"""

print("   测试文档（共11段）:")
for line in test_doc.split('\n')[:3]:
    print(f"      {line}")
print("      ...")

try:
    result = client.analyze_document(test_doc)
    
    if 'paragraphs' in result:
        paragraphs = result['paragraphs']
        print(f"\n   ✅ 识别成功！共识别 {len(paragraphs)} 个段落")
        
        # 统计各类型
        type_counts = {}
        for p in paragraphs:
            ptype = p.get('type', 'unknown')
            type_counts[ptype] = type_counts.get(ptype, 0) + 1
        
        print("\n   📊 识别结果统计:")
        for ptype, count in sorted(type_counts.items()):
            print(f"      {ptype}: {count} 个")
        
        # 显示前3个段落的识别结果
        print("\n   🔍 前3个段落识别详情:")
        for p in paragraphs[:3]:
            idx = p.get('index', '?')
            ptype = p.get('type', 'unknown')
            content = p.get('content', '')[:30]
            print(f"      [{idx}] {ptype}: {content}...")
        
        # 验证识别准确性
        expected_types = {
            0: 'title',
            1: 'recipient',
            3: 'heading1',
            5: 'heading2',
            7: 'heading3',
            9: 'signature',
            10: 'date'
        }
        
        correct = 0
        total = len(expected_types)
        
        print("\n   ✅ 验证关键段落识别:")
        for idx, expected_type in expected_types.items():
            actual = next((p for p in paragraphs if p.get('index') == idx), None)
            if actual:
                actual_type = actual.get('type')
                if actual_type == expected_type:
                    print(f"      [{idx}] ✓ {expected_type}")
                    correct += 1
                else:
                    print(f"      [{idx}] ✗ 期望{expected_type}，实际{actual_type}")
            else:
                print(f"      [{idx}] ✗ 未识别")
        
        accuracy = correct / total * 100
        print(f"\n   📈 准确率: {correct}/{total} = {accuracy:.1f}%")
        
        if accuracy >= 80:
            print("   🎉 识别效果良好！")
        elif accuracy >= 60:
            print("   ⚠️  识别效果一般，可能需要调整Prompt")
        else:
            print("   ❌ 识别效果较差，建议检查模型或Prompt")
    
    else:
        print("   ❌ LLM返回格式错误，缺少paragraphs字段")
        print(f"   返回内容: {result}")
        sys.exit(1)

except Exception as e:
    print(f"   ❌ 识别失败: {str(e)}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

# 测试4: 完整文档处理（可选）
print("\n4️⃣ 完整文档处理测试（可选）")
print("   如需测试完整流程，请准备一个.docx文档")
test_file = input("   输入docx文件路径（直接回车跳过）: ").strip()

if test_file:
    test_file = test_file.strip('"').strip("'").replace('\\ ', ' ')
    
    if os.path.exists(test_file) and test_file.endswith('.docx'):
        print(f"\n   开始处理: {os.path.basename(test_file)}")
        try:
            from llm_formatter import llm_format_document
            success = llm_format_document(test_file)
            if success:
                print("\n   🎉 测试完成！检查生成的llm_xxx.docx文件")
            else:
                print("\n   ❌ 处理失败")
        except Exception as e:
            print(f"\n   ❌ 处理失败: {e}")
    else:
        print("   ⚠️  文件不存在或格式错误")
else:
    print("   ⏭️  跳过完整文档测试")

# 总结
print("\n" + "="*60)
print("  ✅ 所有测试完成！")
print("="*60)
print("\n📝 下一步:")
print("   1. Web模式: python3 app.py")
print("   2. CLI模式: python3 llm_formatter.py")
print("   3. 查看文档: cat README_LLM.md")
print()
