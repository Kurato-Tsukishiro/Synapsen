# Local LLMによるOCR機能 (Experimental)

Synapsen Normalisierer は、ローカルLLM (Ollama) を使用して、従来のTesseract OCRでは読み取りが困難な手書き文字などを認識する実験的な機能を備えています。

本機能は**ファイル単位でのオプトイン**（**明示的な選択**）方式を採用しており、必要なファイルにのみ強力なリソースを割り当てることが可能です。

## 使い方

### 1. 準備
この機能を使用するには、以下の準備が必要です。

1.  **Ollamaのインストール**: [Ollama公式サイト](https://ollama.com/) からインストールし、サーバーを起動しておきます。
2.  **Vision対応モデルの準備**: 画像認識に対応したモデル（例: `llama3.2-vision`, `gemma3`など）をpullしておきます。
    ```bash
    ollama pull llama3.2-vision
    ```
3.  **Pythonパッケージ**: `ollama` ライブラリが必要です。
    - Synapsenのルートフォルダ下で`poetry install --no-root --extras ai` を使用する事で、導入する事が出来ます。
    - 以下のコマンドを使用する事でも導入する事が出来ます。
        ```bash
        poetry add ollama
        ```
        又は
        ```bash
        pip install ollama
        ```

### 2. 設定 (config.ini)
`config.ini` の `[Automation]` セクションに以下の設定を追加します。

```ini
[Automation]

; --- Ollama 設定 ---
; OllamaサーバーのURL (AI_Tagger / Normalisierer 共通)
; ベースURLを指定してください。
ollama_api_url = http://localhost:11434

; --- Ollama (Local LLM) OCR設定 ---
; ファイル名の末尾に "_hand" (または "_llm") が付いている場合、
; Tesseractの設定に関わらず、強制的にOllamaを使用してOCRを行います。
; 画像認識に強いVisionモデルを指定してください
ollama_ocr_model = llama3.2-vision
```

### 3. 実行方法
OCRをかけたい対象のファイル（PDFまたは画像）のファイル名末尾に `_hand` を付与してください。

* 例: `MeetingNote_2024_hand.pdf`
* 例: `IdeaSketch_hand.jpg`

この状態で Normalisierer（フォルダ処理またはD&D）を実行すると、対象ファイルのみOllama経由でOCR処理が行われます。

---

## ⚠️ ベンチマークと推奨環境

本機能は高いハードウェア性能を要求します。
以下は、**ビジネスノートPC (GPUなし)** および **軽量なマルチモーダルモデル** を使用した場合の参考値（非推奨例）です。

### 検証環境 (非推奨スペック)
* **PC**: Panasonic Let's Note CF-SV7
* **CPU**: Intel® Core™ i5-8350U vPro™ (1.70 GHz)
* **GPU**: Intel UHD Graphics 620 (CPU内蔵 / iGPU)
* **RAM**: 8GB
* **使用モデル**: `gemma-3-4b-it-qat-GGUF:Q4_K_M`
    * *注: Gemma 3 (4B) はマルチモーダル対応ですが、パラメータ数が少なく、複雑な手書き文字認識においては精度や安定性が不足する場合があります。*

### 実測結果
* **対象**: 情報カード (A6サイズ相当) に日本語で手書きされた薬理学メモ 1枚
* **処理時間**: 約 3分
* **リソース消費**:
    * CPU使用率: 72% (Ollamaプロセス)
    * メモリ使用量: 約 2.6GB (Ollamaプロセス)

### 認識精度の比較 (軽量モデル・低スペック環境の場合)
軽量モデルでは、専門用語の欠落や幻覚（ハルシネーション）が発生しやすくなります。

| 項目 | 内容 |
| :--- | :--- |
| **原文 (正解)** | **24.7.28 アドレナリン**<br>アドレナリン (Ad)<br>α1作用 > β2作用<br>(血管収縮 => 血圧上昇) (血管拡張 => 血圧下降)<br>↑<br>α1遮断薬併用で血圧下降作用を示す => 血圧反転<br><br>アナフィラキシーショック時に用いるエピペンの主成分<br>↓<br>α1作用による血圧上昇<br>β2作用による気管支拡張 |
| **OCR出力** | **アドレスリン (Acl)**<br>q1作用 > β2作用(血管拡張)↓ 血管拡張(血圧上昇) 血圧下降↓ d.<br>流動薬作用で血圧上昇を元にする↓ 血圧反射<br>3プロドラッグ時に用いるzビパンの主成分 |

### 結論
実用的な速度と精度を求める場合は、**NVIDIA製GPU (VRAM 8GB以上推奨)** を搭載したPCで、**Llama 3.2 Vision (11B)** などの、よりパラメータ数の多いモデルを使用することを強く推奨します。

---

## 利用規約に関する注意 (Models Terms of Use)

使用するモデル（Llama 3.2 Vision, Gemmaシリーズなど）によっては、開発元による利用規約への同意が必要な場合があります。

* **Gemma / Gemma 2 / Gemma 3**:
    Googleによる [Gemma Terms of Use](https://ai.google.dev/gemma/terms) への同意が必要です。Hugging FaceやKaggle等でダウンロードする際、またはOllamaでの初回Pull時に、利用規約を確認し遵守してください。
* **Llama 3.2 Vision**:
    Metaによる [Llama 3.2 Community License Agreement](https://github.com/meta-llama/llama-models/blob/main/models/llama3_2/LICENSE) への同意が必要です。

本ツール (Synapsen) は、ユーザーが指定したモデルをAPI経由で呼び出すインターフェースを提供するのみであり、モデル自体のライセンスや生成内容に関する責任は負いません。各モデルのライセンス条項に従ってご利用ください。