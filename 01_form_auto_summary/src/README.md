# フォーム回答自動要約レポート

Googleフォームの回答を自動取得し、ChatGPT APIで要約 → Google Docs → PDF → Gmail送信するツール。

## 使い方
1. スプレッドシートを用意し「FormResponses」シートを作成
2. Apps ScriptでこのCode.gsを貼り付け
3. OPENAI_API_KEYをScript Propertiesに登録
4. seedFormResponsesDemo() → summarizeFormResponsesAndSend() を実行