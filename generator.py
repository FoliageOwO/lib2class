from typing import List, Tuple
from docx import Document as get_document
from docx.document import Document
import os, re, json

docs_dir = "docs"

docs = [
    name
    for name in os.listdir(docs_dir)
    if name.endswith(".doc") or name.endswith(".docx")
]
print(f"检测到 {len(docs)} 个文档")


def get_questions_and_answers(doc: Document) -> List[Tuple[str, str]]:
    result = []
    question = None
    options = {}
    for text in [p.text for p in doc.paragraphs]:
        text = text.strip()

        question_match = re.match("^(\d+[.。,，、:：．])(.*)", text)
        option_match = re.match("^([A-Z])([.。,，、．])(.*)", text)
        answer_match = re.match("^(答案[:：])([A-Z])", text)

        if question_match:
            question_groups = question_match.groups()
            if len(question_groups) == 2:
                question = question_groups[1]

        if option_match:
            option_groups = option_match.groups()
            if len(option_groups) == 3:
                option = option_groups[0]
                content = option_groups[2]
                options[option] = content

        if answer_match:
            answer_groups = answer_match.groups()
            if len(answer_groups) == 2:
                answer = answer_groups[1]
                result.append((question, options[answer]))
                question = None
                options = {}

    return result


for doc_name in docs:
    file = get_document(f"{docs_dir}/{doc_name}")
    result = []
    qnas = get_questions_and_answers(file)
    print(f"读取了题库文件 `{doc_name}` 共 {len(qnas)} 个题目")
    for qna in qnas:
        result.append({"question": qna[0], "answer": qna[1]})
    data = json.dumps(result, ensure_ascii=False, indent=2)

    js_file = open(doc_name.strip(".docx").strip(".doc") + ".json", "w+").write(data)
