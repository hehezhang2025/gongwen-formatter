#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Ollama 客户端 - 调用本地 Qwen 模型
"""

import requests
import json
from config import OLLAMA_CONFIG


class OllamaClient:
    """Ollama 本地大模型客户端"""
    
    def __init__(self, base_url=None, model=None):
        self.base_url = base_url or OLLAMA_CONFIG["base_url"]
        self.model = model or OLLAMA_CONFIG["model"]
        self.temperature = OLLAMA_CONFIG["temperature"]
        self.timeout = OLLAMA_CONFIG["timeout"]
    
    def check_connection(self):
        """检查 Ollama 是否运行"""
        try:
            response = requests.get(f"{self.base_url}/api/tags", timeout=5)
            if response.status_code == 200:
                models = response.json().get("models", [])
                model_names = [m["name"] for m in models]
                
                # 检查目标模型是否存在
                if self.model in model_names:
                    return True, f"✅ Ollama运行正常，找到模型: {self.model}"
                else:
                    return False, f"❌ 模型 {self.model} 不存在。可用模型: {', '.join(model_names)}"
            else:
                return False, f"❌ Ollama响应异常: {response.status_code}"
        except requests.exceptions.ConnectionError:
            return False, f"❌ 无法连接到 Ollama ({self.base_url})，请确保 Ollama 已启动"
        except Exception as e:
            return False, f"❌ 连接检查失败: {str(e)}"
    
    def analyze_document(self, document_text):
        """调用 Qwen 模型分析文档结构"""
        prompt = self._build_prompt(document_text)
        
        try:
            response = requests.post(
                f"{self.base_url}/api/generate",
                json={
                    "model": self.model,
                    "prompt": prompt,
                    "stream": False,
                    "temperature": self.temperature,
                    "options": {
                        "temperature": self.temperature,
                        "num_predict": 4096  # 最大输出token数
                    }
                },
                timeout=self.timeout
            )
            
            if response.status_code != 200:
                raise Exception(f"Ollama API 调用失败: HTTP {response.status_code}")
            
            result = response.json()
            response_text = result.get("response", "")
            
            if not response_text:
                raise Exception("Ollama 返回空结果")
            
            # 解析 JSON 响应
            try:
                # 尝试提取 JSON（可能被包裹在其他文字中）
                import re
                json_match = re.search(r'\{.*\}', response_text, re.DOTALL)
                if json_match:
                    json_str = json_match.group(0)
                    parsed_result = json.loads(json_str)
                else:
                    parsed_result = json.loads(response_text)
                
                return parsed_result
                
            except json.JSONDecodeError as e:
                raise Exception(f"无法解析 LLM 返回的 JSON: {str(e)}\n返回内容: {response_text[:500]}")
        
        except requests.exceptions.Timeout:
            raise Exception(f"LLM 调用超时（超过 {self.timeout} 秒）")
        except requests.exceptions.ConnectionError:
            raise Exception("无法连接到 Ollama，请确保服务已启动")
        except Exception as e:
            raise Exception(f"LLM 调用失败: {str(e)}")
    
    def _build_prompt(self, document_text):
        """构建用于文档结构识别的 Prompt"""
        return f"""你是公文结构识别专家。请严格按照GB/T 9704-2012标准分析以下文档，识别每个段落的类型。

【识别规则】
1. title（标题）: 包含"通知"、"报告"、"决定"、"意见"、"办法"、"方案"等文种词，通常在前3段
2. recipient（主送机关）: 以"："或":"结尾，包含"局"、"委"、"厅"、"部"、"各"等关键词
3. heading1（一级标题）: "一、"、"二、"、"三、"开头，或包含关键动词的6-20字短语（如"加强XX"、"推进XX"）
4. heading2（二级标题）: "（一）"、"（二）"、"（三）"开头
5. heading3（三级标题）: "1."、"2."、"3."开头（注意是半角点号）
6. heading4（四级标题）: "(1)"、"(2)"、"(3)"开头（注意是半角括号）
7. body（正文）: 普通段落，以"为"、"根据"、"按照"、"经"等开头，或正常叙述性文字
8. attachment_marker（附件标记）: "附件："或"附件1："等，单独一行
9. signature（署名）: 包含单位名称，位于文档后部，通常在日期前一行
10. date（日期）: 包含"年月日"格式，位于文档末尾

【重要规则】
- 附件标记后的内容，标题编号会重新开始
- 如果一个段落同时符合多个特征，优先选择更具体的类型（如标题>正文）
- 不确定时标记为body（正文）
- 表格、图片说明标记为body

【输出格式要求】
严格按以下JSON格式输出，不要包含任何其他文字或解释：
{{
  "paragraphs": [
    {{"index": 0, "type": "title", "content": "段落内容"}},
    {{"index": 1, "type": "recipient", "content": "段落内容"}},
    {{"index": 2, "type": "body", "content": "段落内容"}},
    {{"index": 3, "type": "heading1", "content": "段落内容"}},
    ...
  ],
  "attachment_start_index": 25
}}

注意：
- index 必须与下面文档中的行号一致
- type 只能是上述10种类型之一
- content 必须与原文一致
- attachment_start_index 是附件标记所在的index，如果没有附件则设为-1

【文档内容】（行号: 内容）
{document_text}

请开始分析，只输出JSON，不要任何额外文字："""


def test_ollama_connection():
    """测试 Ollama 连接（供调试使用）"""
    print("\n🔍 测试 Ollama 连接...")
    print(f"   地址: {OLLAMA_CONFIG['base_url']}")
    print(f"   模型: {OLLAMA_CONFIG['model']}")
    
    client = OllamaClient()
    success, message = client.check_connection()
    print(f"   {message}")
    
    if success:
        # 测试一个简单的文档识别
        print("\n🧪 测试文档识别...")
        test_doc = """0: 关于加强项目管理的通知
1: 各部门：
2: 为了提高项目管理水平，现通知如下。
3: 一、加强组织领导
4: 各部门要高度重视。
5: （一）成立工作小组
6: 由部门负责人担任组长。
7: XX公司
8: 2025年2月17日"""
        
        try:
            result = client.analyze_document(test_doc)
            print("   ✅ 识别成功！")
            print(f"   识别到 {len(result.get('paragraphs', []))} 个段落")
            return True
        except Exception as e:
            print(f"   ❌ 识别失败: {str(e)}")
            return False
    
    return False


if __name__ == '__main__':
    # 运行测试
    test_ollama_connection()
