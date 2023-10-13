import os
import json
import requests
import pandas as pd
import hashlib
import random
from openpyxl import load_workbook

# 常量: 你的百度翻译 API credentials
BAIDU_URL = "https://fanyi-api.baidu.com/api/trans/vip/translate"
BAIDU_APP_ID = "*******"  # 替换为你的百度应用ID
BAIDU_SECRET_KEY = "*****"  # 替换为你的百度应用密钥


def extract_words_from_excel(file_path):
    """从 Excel 文件中提取单词并返回一个列表"""
    df = pd.read_excel(file_path)
    # 假设单词位于第一列
    words = df.iloc[:, 0].dropna().tolist()
    return words


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
                            print(
                                f"Warning: Missing 'trans' key for word '{word_data['name']}' in file {file_name}. Full data: {word_data}")
                            vocab[word_data["name"]] = ""
    return vocab


def baidu_translate(word):
    """使用百度翻译API进行翻译"""
    q = word
    from_lang = 'en'
    to_lang = 'zh'
    salt = str(random.randint(32768, 65536))
    sign = BAIDU_APP_ID + q + salt + BAIDU_SECRET_KEY
    sign = hashlib.md5(sign.encode()).hexdigest()
    params = {
        'q': q,
        'from': from_lang,
        'to': to_lang,
        'appid': BAIDU_APP_ID,
        'salt': salt,
        'sign': sign,
    }
    response = requests.get(BAIDU_URL, params=params)
    if response.status_code == 200:
        data = response.json()
        translations = data.get("trans_result", [])
        if translations:
            return translations[0]['dst']
    return f"Translation Error (HTTP Status: {response.status_code})"


def local_translate(word, vocab):
    translation = vocab.get(word)
    if translation:
        return translation
    else:
        # 使用百度翻译API进行翻译
        return baidu_translate(word)


def main():
    VOCAB_FOLDER_PATH = "/Users/loyo/PycharmProjects/pythonProject2/1_1_中文释义"
    vocab = load_vocabulary_from_folder(VOCAB_FOLDER_PATH)

    file_path = "*****"  # 更新为你的 Excel 文件路径
    words = extract_words_from_excel(file_path)  # 更新函数调用

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