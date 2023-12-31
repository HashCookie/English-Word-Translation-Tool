# 英文单词翻译工具

这是一个 Python 工具，它从 Word 文档中提取英文单词，并使用本地词汇表或有道词典 API 将它们翻译成中文。

## 主要功能

- 从 Word 文档中提取英文单词。
- 使用本地 JSON 词汇表作为首选的翻译来源。
- 如果本地词汇表中不存在某个单词的翻译，那么会使用有道词典 API 进行翻译。

## 如何使用

1. 确保你的机器上已安装了 Python 3 和所有必要的库，例如 `docx` 和 `requests`。
2. 修改 `VOCAB_FOLDER_PATH` 和 `file_path`，使它们指向正确的文件或文件夹路径。
3. 运行 `python filename.py`（请将 `filename.py` 替换为您的 Python 文件名）。

## 注意事项

请确保您的有道词典 API 密钥是有效的，并注意不要过于频繁地调用API，以避免可能的封禁。

## 许可

此项目采用 MIT 许可证。
