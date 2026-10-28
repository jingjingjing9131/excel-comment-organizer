# 表格整理台

本地运行的 Excel 整理工具：选择模板工作表和评论来源工作表，使用 DeepSeek 拆分评论并复核评论行数、评论编号/原文列位置及模板表头映射，在网页中编辑结果，再导出沿用模板单元格样式的 `.xlsx` 工作簿。每条非空原文行都单独作为一条评论；如果模型试图合并、漏掉或重排原文行，该篇笔记会回退到本地逐行拆分。样式由程序直接复制模板，DeepSeek 复核的是表头和数据结构。

## 启动

需要 Python 3.10 或更新版本。

```bash
cd excel-assistant
python3 -m venv .venv
source .venv/bin/activate
python3 -m pip install -r requirements.txt
cp .env.example .env
```

编辑 `.env`，填入你自己的 DeepSeek API Key：

```text
DEEPSEEK_API_KEY=你的密钥
DEEPSEEK_MODEL=deepseek-flash
```

然后启动：

```bash
python3 app.py
```

浏览器打开 <http://127.0.0.1:8765>。按 `Ctrl+C` 停止服务。

未配置 Key 时仍可用本地规则拆分常见的“单评/盖楼”格式，并能在网页中手动编辑；AI 格式复核会显示为跳过。使用 DeepSeek 时，拆分请求发送笔记标题和评论原文；格式复核请求发送模板表头、列映射和拆分后的评论编号/原文。密钥只由本地后端读取，不会发送到浏览器。

## 运行检查

```bash
python3 -m unittest discover -s tests -v
```
