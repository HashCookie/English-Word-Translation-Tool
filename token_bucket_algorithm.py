import requests
import hashlib
import time
import random
import json
from docx import Document

APP_KEY = '*****'
APP_SECRET = '******'
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

class TokenBucket:
    def __init__(self, capacity, refill_rate):
        self.capacity = capacity
        self._tokens = capacity
        self.refill_rate = refill_rate
        self.last_refill = time.time()

    def _refill(self):
        now = time.time()
        since_last = now - self.last_refill
        new_tokens = since_last * self.refill_rate
        self._tokens = min(self.capacity, self._tokens + new_tokens)
        self.last_refill = now

    def consume(self):
        if self._tokens < 1:
            self._refill()
        if self._tokens < 1:
            return False
        self._tokens -= 1
        return True

def main():
    file_path = "*****"
    words = extract_words_from_docx(file_path)

    translated_words = []
    total_words = len(words)

    # 初始化令牌桶，假设我们的容量是10个令牌，每秒填充1个令牌
    bucket = TokenBucket(10, 1)

    for idx, word in enumerate(words, 1):
        # 等待直到我们可以消耗一个令牌
        while not bucket.consume():
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