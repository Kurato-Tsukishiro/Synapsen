import sqlite3
import configparser
import json
import time
import re
from ollama import Client

# --- 設定 ---
# タイムアウト対策: Client側でタイムアウトを長めに設定する(必要であれば)

# システムプロンプト（役割定義のみ）
SYSTEM_PROMPT = """
あなたはナレッジベースの整理アシスタントです。
ユーザーから提供されたテキストから「検索用キーワード（5個程度）」と「3行程度の要約」を作成してください。
"""


def clean_json_text(text):
    """
    LLMが余計なMarkdown記号（```json ... ```）をつけた場合に削除する
    """
    cleaned = re.sub(r"^```json\s*", "", text, flags=re.MULTILINE)
    cleaned = re.sub(r"^```\s*", "", cleaned, flags=re.MULTILINE)
    cleaned = re.sub(r"\s*```$", "", cleaned, flags=re.MULTILINE)
    return cleaned.strip()


def get_ai_metadata(client, text, model):
    """Ollama ライブラリを使用してメタデータを取得"""
    if not text or len(text) < 10:
        return None

    # テキストが長すぎる場合の安全策 (モデルのコンテキストウィンドウに合わせる)
    # 日本語はトークン数が増えがちなので、安全を見て3000文字程度でカット
    prompt_text = text[:3000]

    # 【重要】小規模モデル対策：指示を入力テキストの「後」に配置する
    user_message = f"""
以下のテキストを分析してください。

---
{prompt_text}
---

【指示】
上記のテキストの「重要なキーワード」と「要約」を抽出し、以下のJSONフォーマット**のみ**を出力してください。
挨拶や解説は不要です。

{{
  "keywords": ["単語1", "単語2", "関連語"],
  "summary": "要約文..."
}}
"""

    try:
        response = client.chat(
            model=model,
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": user_message},
            ],
            format="json",  # OllamaのJSONモードを強制
            options={
                "num_ctx": 4096,  # コンテキストウィンドウを拡張
                "temperature": 0.1,  # 創造性を下げてフォーマット遵守率を上げる
            },
        )

        content = response["message"]["content"]

        # JSONとしてパースできるか試みる
        try:
            cleaned_content = clean_json_text(content)
            return json.loads(cleaned_content)
        except json.JSONDecodeError:
            print("  [Warning] JSON Decode failed. Raw output returned.")
            # JSONパース失敗時は、生のテキストをsummaryとして無理やり返す
            return {"keywords": ["ParseError"], "summary": content}

    except Exception as e:
        print(f"  [AI Error] {e}")
        return None


def main():
    config = configparser.ConfigParser()

    # 読み込み (カレントディレクトリ または 2階層上(ルート)を探すようにすると親切です)
    read_files = config.read(["config.ini", "../../config.ini"], encoding="utf-8")

    if not read_files:
        print(
            "config.iniが見つかりません。デフォルト設定で続行、もしくはパスを確認してください。"
        )
        # config.iniがない場合の完全なフォールバック
        db_path = "Synapsen.db"
        api_url = "http://localhost:11434"
        model = "hf.co/unsloth/gemma-3-4b-it-qat-GGUF:Q4_K_M"
    else:

        # db_path: 必須に近いですが、なければデフォルトを使用
        # config.iniがあっても [Paths] セクションがない場合のエラーハンドリングも兼ねる
        db_path = config.get("Paths", "database_path", fallback="Synapsen.db")

        # api_url: 設定がなければデフォルト値を使用
        api_url = config.get(
            "Automation", "ollama_api_url", fallback="http://localhost:11434"
        )

        # model: 設定がなければデフォルト値を使用
        model = config.get(
            "Automation",
            "ollama_tagger_model",
            fallback="hf.co/unsloth/gemma-3-4b-it-qat-GGUF:Q4_K_M",
        )

    # Ollamaクライアントの初期化
    try:
        client = Client(host=api_url)
    except Exception as e:
        print(f"Ollamaへの接続に失敗しました: {e}")
        return

    print(f"データベース: {db_path}")
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    # 1. 処理対象の検索
    print("データベースから未処理のノートを検索中...")
    cursor.execute(
        """
        SELECT key, title, full_text, memo
        FROM notes
        WHERE full_text IS NOT NULL
          AND full_text != ''
          AND (memo IS NULL OR memo NOT LIKE '%[AI-Summary]%')
          AND (tags IS NULL OR tags NOT LIKE '%SkipAISummary%')
    """
    )
    notes = [dict(row) for row in cursor.fetchall()]

    print(f"処理対象: {len(notes)} 件")

    if len(notes) == 0:
        print("未処理のノートはありません。")
        conn.close()
        return

    proceed = input(f"モデル '{model}' を使用して処理を開始しますか？ (y/n): ")
    if proceed.lower() != "y":
        conn.close()
        return

    processed_count = 0

    try:
        for note in notes:
            title_display = (
                (note["title"][:30] + "..")
                if len(note["title"]) > 30
                else note["title"]
            )
            print(
                f"[{processed_count+1}/{len(notes)}] Processing: {title_display} ... ",
                end="",
                flush=True,
            )

            start_time = time.time()
            ai_data = get_ai_metadata(client, note["full_text"], model)
            elapsed = time.time() - start_time

            if ai_data:
                keywords = ", ".join(ai_data.get("keywords", []))
                summary = ai_data.get("summary", "")

                # --- DB更新処理 ---
                original_memo = note["memo"] if note["memo"] else ""

                # 既存のAIタグがあれば重複を避ける等の処理が必要ですが、
                # ここではSQLでフィルタリングしているので単純追記します
                append_text = f"\n\n[AI-Summary]\n{summary}\n[AI-Keywords]\n{keywords}"
                new_memo = original_memo + append_text

                cursor.execute(
                    "UPDATE notes SET memo = ? WHERE key = ?", (new_memo, note["key"])
                )
                conn.commit()

                print(f"Done ({elapsed:.1f}s)")
                processed_count += 1
            else:
                print("Skipped (No Data)")

            # マシン負荷を考慮して少し待機
            # 連続して重い処理を行うとPC全体の動作が重くなるため
            time.sleep(1.0)

    except KeyboardInterrupt:
        print("\n処理を中断しました。")
    finally:
        conn.close()
        print(f"完了しました。 {processed_count} 件のノートを更新しました。")


if __name__ == "__main__":
    main()
