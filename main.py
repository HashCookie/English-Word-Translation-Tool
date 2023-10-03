import os
import json
import requests
from docx import Document

# 常量: 你的有道词典 API credentials
YOUDAO_URL = "http://openapi.youdao.com/api"
YOUDAO_KEY = "此处替换你的YOUDAO_KEY"
YOUDAO_KEYFROM = "此处替换你的YOUDAO_KEYFROM"

def extract_words_from_docx(file_path):
    """从 Word 文档中提取单词并返回一个列表"""
    doc = Document(file_path)
    return [p.text for p in doc.paragraphs if p.text.strip() != '']

def load_vocabulary_from_folder(folder_path):
    vocab = {}
    for root, dirs, files in os.walk(folder_path):
        for file_name in files:
            if file_name.endswith('.json'):
                file_path = os.path.join(root, file_name)
                with open(file_path, 'r') as f:
                    words = json.load(f)
                    for word_data in words:
                        if "name" not in word_data:
                            print(f"Warning: Missing 'name' key in file {file_name}. Full data: {word_data}")
                            continue
                        if "trans" in word_data:
                            vocab[word_data["name"]] = word_data["trans"][0]
                        else:
                            print(f"Warning: Missing 'trans' key for word '{word_data['name']}' in file {file_name}. Full data: {word_data}")
                            vocab[word_data["name"]] = ""
    return vocab

def youdao_translate(word):
    """使用有道词典API进行翻译"""
    params = {
        "q": word,
        "from": "EN",
        "to": "zh-CHS",
        "key": YOUDAO_KEY,
        "keyfrom": YOUDAO_KEYFROM,
        "type": "data",
        "doctype": "json",
        "version": "1.1"
    }
    response = requests.get(YOUDAO_URL, params=params)
    if response.status_code == 200:
        data = response.json()
        translations = data.get("translation", [])
        if translations:
            return translations[0]
        # 如果返回的 JSON 中有 errorCode 字段，打印它
        error_code = data.get("errorCode")
        if error_code:
            print(f"Error code from Youdao for word '{word}': {error_code}")
            return f"Translation Error (Code: {error_code})"
    return f"Translation Error (HTTP Status: {response.status_code})"

def local_translate(word, vocab):
    translation = vocab.get(word)
    if translation:
        return translation
    else:
        # 使用有道词典 API 进行翻译
        return youdao_translate(word)

def main():
    VOCAB_FOLDER_PATH = "此处替换本地词库"
    vocab = load_vocabulary_from_folder(VOCAB_FOLDER_PATH)

    file_path = "此处替换文件地址"
    words = extract_words_from_docx(file_path)

    translated_words = []
    total_words = len(words)
    for idx, word in enumerate(words, 1):
        translation = local_translate(word, vocab)
        translated_words.append({
            "name": word,
            "trans": [translation]
        })
        print(f"Translated {idx}/{total_words}: {word} -> {translation}")

    with open("translated_words.json", "w") as f:
        json.dump(translated_words, f, ensure_ascii=False, indent=4)

if __name__ == "__main__":
    main()