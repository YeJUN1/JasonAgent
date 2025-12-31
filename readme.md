在调用前，请确保已安装火山引擎SDK，并将该SDK升级到最新版本更多信息请查看SDK指南：
https://www.volcengine.com/docs/82379/1541595?lang=zh

Doubao API 接入示例：
1) 设置环境变量：`export ARK_API_KEY=你的API密钥`
2) 运行示例：
   - `python Src/doubao_client.py`
3) 在代码里调用：
   - `from doubao_client import build_messages, chat_completion`

通用文字识别 OCR 配置（视觉服务）：
- API 文档：https://www.volcengine.com/docs/86081/1660261?lang=zh
- 在 `.env` 中设置 `VOLC_ACCESS_KEY`、`VOLC_SECRET_KEY`
- 影印版 PDF 与图片均使用该 OCR 配置
- 可选参数：
  - `OCR_REGION`（默认 `cn-north-1`）
  - `OCR_ENDPOINT`（默认 `visual.volcengineapi.com`）
  - `OCR_IMAGE_MODE`（`base64`/`image_url`，默认 `base64`）
  - `OCR_IMAGE_URL_PREFIX`（当 `OCR_IMAGE_MODE=image_url` 时必填）
  - `OCR_MODE`（`default`/`text_block`）
  - `OCR_FILTER_THRESH`、`OCR_APPROXIMATE_PIXEL`、`OCR_HALF_TO_FULL`
  - `OCR_MAX_WORKERS`（并发识别线程数，默认 2-4）
