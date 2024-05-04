"""
从 Word 文档中提取所有的 JSON 对象
"""

from docx import Document
import json

def extract_json_from_docx(doc_path):
    """从 Word 文档中提取所有的 JSON 对象"""
    document = Document(doc_path)
    json_objects = []

    all_text = ""
    for para in document.paragraphs:
        all_text += para.text.strip()

    stack = []
    start_index = 0
    for i in range(len(all_text)):
        if all_text[i] == "{":
            stack.append("{")
            if len(stack) == 1:  # 当栈中只有一个 "{" 时，记录开始索引
                start_index = i
        elif all_text[i] == "}":
            if stack:  # 如果栈不为空
                stack.pop()
                if not stack:  # 如果弹出后栈为空，表示找到一个完整的 JSON 对象
                    json_str = all_text[start_index: i+1]
                    try:
                        json_obj = json.loads(json_str)
                        json_objects.append(json_obj)
                        print(f"有效 JSON: {json_obj}")
                    except json.JSONDecodeError:
                        print(f"无效 JSON: {json_str}")

    return json_objects