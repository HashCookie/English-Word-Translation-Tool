import requests
import hashlib
import time
import random
import json
from docx import Document

APP_KEY = '********'
APP_SECRET = '*******'
YOUDAO_URL = 'https://openapi.youdao.com/api'

def encrypt(signStr):
    hash_algorithm = hashlib.sha256()
    hash_algorithm.update(signStr.encode('utf-8'))
    return hash_algorithm.hexdigest()

def do_request(data):
    headers = {
        'Content-Type': 'application/x-www-form-urlencoded',
    }
    response = requests.post(YOUDAO_URL, data=data, headers=headers)
    return response.content

def translate_to_chinese(query):
    try:
        data = {}
        data['from'] = 'EN'
        data['to'] = 'zh-CHS'
        data['signType'] = 'v3'
        curtime = str(int(time.time()))
        data['curtime'] = curtime
        salt = str(random.randint(0, 10))
        signStr = APP_KEY + truncate(query) + salt + curtime + APP_SECRET
        sign = encrypt(signStr)
        data['appKey'] = APP_KEY
        data['q'] = query
        data['salt'] = salt
        data['sign'] = sign
        data['vocabId'] = "您的词库ID"

        response = do_request(data)
        response_data = json.loads(response)

        if 'translation' in response_data:
            return response_data['translation'][0]
        else:
            print(f"Warning: No translation found for word '{query}'")
            return ""
    except Exception as e:
        print(f"Error translating word '{query}': {e}")
        return ""

def truncate(q):
    size = len(q)
    return q if size <= 20 else q[0:10] + str(size) + q[size - 10:size]

def extract_words_from_docx(file_path):
    doc = Document(file_path)
    return [p.text for p in doc.paragraphs if p.text.strip() != '']

class LeakyBucket:
    def __init__(self, capacity, leak_rate):
        self.capacity = capacity
        self._water = 0
        self.leak_rate = leak_rate
        self.last_leak_time = time.time()

    def _leak_out(self):
        now = time.time()
        since_last = now - self.last_leak_time
        leak_amount = since_last * self.leak_rate
        self._water = max(0, self._water - leak_amount)
        self.last_leak_time = now

    def can_consume(self):
        self._leak_out()
        if self._water < self.capacity:
            self._water += 1
            return True
        else:
            return False

def main():
    file_path = "******"
    words = extract_words_from_docx(file_path)

    translated_words = []
    total_words = len(words)

    # 初始化漏桶，假设容量为10，每秒泄漏1个请求
    bucket = LeakyBucket(10, 1)

    for idx, word in enumerate(words, 1):
        # 等待直到我们可以消耗一个请求
        while not bucket.can_consume():
            time.sleep(0.1)

        translation = translate_to_chinese(word)
        translated_words.append({
            "name": word,
            "trans": [translation]
        })
        print(f"Translated {idx}/{total_words}: {word} -> {translation}")

    with open("translated_words.json", "w") as f:
        json.dump(translated_words, f, ensure_ascii=False, indent=4)

if __name__ == '__main__':
    main()