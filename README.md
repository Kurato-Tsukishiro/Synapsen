![banner](assets/banner_source/banner.svg)

# Synapsen - PDFノート管理ツールセット

`Synapsen` (シナプセン) は、スキャンしたPDFノートやドキュメントを、デジタル・ツェッテルカステン（Zettelkasten）風に管理・閲覧するためのツールセットです。

このプロジェクトは、以下の3つの独立したアプリケーションで構成されています。

1.  **Normalisierer (正規化):** PDFのフォームをテキスト化し、指定サイズ (A4/A5) に統一します。
2.  **Ersteller (作成・統合):** ノートにメタデータを付与・抽出し、月ごとに1冊のPDFに統合します。
3.  **Nexus (閲覧・検索):** 統合されたノートのデータベースを、強力な検索・リンク機能で閲覧・編集します。

## 思想・コンセプト

このツールセットは、書籍[『情報は1冊のノートにまとめなさい［完全版］』](https://ndlsearch.ndl.go.jp/books/R100000002-I025014527)(奥野 宣之 著, ISBN: 9784478022009) で紹介されている、**情報を一元管理する**という思想に強く感銘を受けて作成されました。アナログ・デジタルの情報を一元化し、「時系列」で管理しつつ「タグ」や「目次」で検索性を高めるという理論を、スキャンPDFや電子ペーパー(デジタルノート)で実現することを目的としています。

これに加え、本ツールは伝統的な知識管理術である「**コモンプレイス・ブック (Commonplace Book)**」の概念も取り入れています。
本ツール独自の **Index Key (コモンプレイス Key)** は、この「コモンプレイス」のKey（索引）の概念に由来しており、時系列やタグとは異なる「テーマ」や「概念」でノートを横断するための機能です。

## 構成ツールと機能

### 1. Synapsen Normalisierer (正規化ツール)

スキャンしたPDF、Markdownファイル、Webページ、画像などを、`Ersteller` で処理できる形式に変換します。

* **PDFフォームのフラット化:** PDFフォームの入力内容を、注釈（アノテーション）を維持したままテキストに変換（フラット化）します。
* **サイズ正規化:** すべてのPDFページを `config.ini` で指定された用紙サイズ（A4またはA5）の縦サイズに（アスペクト比を維持して）リサイズ・中央配置します。
* **D&D / ペーストによるクリップ:**
    * `PDF`, `PNG`, `JPG`, `MD` (Markdown) ファイルのD&D、またはクリップボードのスクリーンショット（Ctrl+V）に対応します。
    * 実行時に **IndexKey**, **コメント**, **書誌情報** を一括で指定可能です。
    * **Markdown (.md) ファイル**は、`Pandoc`と`Playwright`を経由してPDFに自動変換されます。`<details>`タグは自動的に展開（`<details open>`）され、内容がPDFに含まれます。
* **Webクリップ (URLからPDF化):**
    * 指定したURLを `Playwright` を使用してPDFとして保存します。
    * **PDFや画像ファイルへの直接URL**にも対応。HTML以外はファイルを直接ダウンロードして正規化します。
    * 実行時に **IndexKey**, **コメント**, **書誌情報**（著者名、サイト名など）を入力可能です。
* **クリップ形式の統一:**
    * 「D&D/ペースト」および「Webクリップ」で生成されるPDFは、**1ページ目にIndexKey**が埋め込まれ、**最終ページ（新規追加）にコメントと書誌情報**が挿入される形式に統一されます。
* **テキスト抽出 (OCR):**
    * PDFに埋め込まれた**既存のテキストレイヤー（テキストベースPDF）を抽出**します。
    * (オプション) `config.ini` で `enable_tesseract_ocr = true` に設定すると、テキストレイヤーが存在しない**画像PDF**（画像クリップやスキャンPDF）に対し、Tesseract OCR を使用してテキストを抽出し、検索可能な「透明テキストレイヤー」としてPDFに埋め込みます。

### 2. Synapsen Ersteller (統合・作成ツール)

正規化されたPDFを読み込み、メタデータを編集し、月報のような形で1つのPDFにまとめ上げます。

* ファイル名から日付やタイトルを自動抽出 (`YYYYMMDD_hhmmss_タイトル.pdf` 形式を推奨)。
    * **（重要）**
        > * この形式は `Nexus` でのリンク機能（`[[key]]`）の基盤となるユニークID（`key`）を生成するために必須です。
        > * この形式でないファイルは `key` を持たず、マスターDBへの登録やリンク対象から除外されます。
* PDFの特定座標からIndex Keyを自動抽出（`config.ini` の `[Extraction]` で設定）。
* "サイドノート"（例: `..._Note.pdf`）に親ノートのIndex Keyを自動継承。
    * これは 電子ペーパー「QUADERNO（クアデルノ）」の "サイドノート" 機能を意識した物です。
* ノートごとにタグ、メモ、Index Key（索引キー）を編集。
* 複数のノートを選択し、Index Keyやタグを一括で設定（追加・削除）する**バッチ編集機能**。
* 指定した月のノート群を1つのPDFに統合。
* （LuaLaTeXを使用し）目次、タグ索引、Index Key索引を自動生成。
* **データ保全・復旧機能:**
    * **メタデータ埋め込み:** 統合PDF生成時、全ノートのメタデータ（JSON）をPDFの添付ファイルとして埋め込みます。これにより、元ファイルやDBが消失しても**統合PDF単体から完全に復旧可能**です。
    * **DB復旧ツール:** 万が一データベースが破損・紛失した場合でも、統合PDFからデータベースを再構築できる専用GUIツールを搭載しています。
      * ウィンドウ右上のボタンから起動します。
      * 初期状態では表示が隠れているため、使用する際はウィンドウの横幅を広げてください。
    * **容量削減:** 復旧時、本文テキスト（`full_text`）は統合PDFのページから再抽出するため、埋め込みデータのサイズは最小限に抑えられます。
* 「統合PDFを生成」する際、`Normalisierer` が埋め込んだ**本文テキスト（`full_text`）をPDFからオンデマンドで抽出し**、マスターDBに保存。
* 統合PDFの索引情報となる**マスターDB（SQLite / `.db`ファイル）**（`Nexus`が使用）に**情報を追記**。
* 現在の作業リストを中間ファイル（CSV）として（`full_text` を除いて）保存・読み込みする機能。

### 3. Synapsen Nexus (閲覧・検索ツール)

`Ersteller` が作成・追記した**マスターDB（SQLite / `.db`ファイル）**を読み込む、ノート閲覧・検索用のアプリケーションです。

* **高度な検索:**
    * **非同期検索:** 検索ロジックをバックグラウンドで実行することで、大量のノートや重い全文検索を行ってもアプリの動作がフリーズしません。
    * `AND`, `OR`, `NOT(-)`, `( )` 演算子を使った検索。
    * `tag:`, `memo:`, `ikey:`, `fulltext:` (または `text:`) などのプレフィックス検索。
    * `date:>=YYYYMMDD` (以降), `date:<=YYYYMMDD` (以前), `date:YYYYMMDD-YYYYMMDD` (期間) などの**高度な日付範囲検索**。
    * PDF本文（`full_text`）を検索対象に含める「**本文検索**」チェックボックス。
    * 「保存済み検索」でよく使う検索を登録可能。
* **プレビューと編集:**
    * 検索結果のノートをダブルクリックすると、統合PDFの該当ページを直接表示。
    * 検索結果のノートをシングルクリックすると、右ペインに詳細と**PDFの1ページ目インラインプレビュー**を表示。
    * メモ内の `[[key]]` 形式のリンクから、別のノートを**簡易プレビュー**（読み取り専用ウィンドウ）で表示。**プレビューウィンドウから直接編集画面へ遷移**することも可能。
    * ノート詳細表示時に、そのノートを引用している他のノート（被リンク元）を自動でリストアップ。
* **グラフ可視化:**
    * **全体グラフ (Global):** 現在の検索結果に基づき、ノート間のリンクを全体的に可視化。
    * **関連グラフ (Local):** 選択したノートを中心とした、直接的なつながりを可視化。
    * **選択グラフ (Selected):** 複数選択したノートのみの関係性を可視化し、一時的な比較や構造化に利用可能。
    * グラフ上のノードをダブルクリックでPDFを開いたり、右クリックでKeyをコピーすることが可能。
* **ナレッジワーク(創造的な情報の活用)支援機能:**
    * **複数選択:** チェックボックスによるノートの複数選択。
    * **リンクの一括コピー:** 選択したノートのリンク（`[[Key: Title]]`）を一括でクリップボードにコピーし、MOC（目次ノート）の作成を支援。
    * **ランダムノート（閃き）:** 現在の検索結果や全ノートからランダムでノートを表示。
    * **保存済み検索:** 現在の検索クエリを保存・管理・呼び出しできる「スマートフォルダ」機能。
* **データ保護:**
    * **自動バックアップ:** アプリ終了時に、データベースのバックアップコピーを `db_backups` フォルダに自動作成します（日次更新）。
* **エクスポート:**
    * 検索結果全体、または**選択したノートのみ**を対象にエクスポート可能。
    * メタデータCSV、本文TXT、グラフHTMLに加え、対象ノートを結合した**統合PDF（しおり付き）**を一括で出力可能。
* **DB直接編集:**
    * ノートのメタデータ（メモ、タグ、Index Key）をDBに直接書き込み・編集・削除する機能。

### 共通機能
* **ロギング:** 全ツールにおいて、エラー発生時や処理状況を `logs/` フォルダへ自動的に記録します。これにより、トラブルシューティングが容易になります。

## 動作環境・依存関係

* **Python 3.x**
* **LuaLaTeX** (TeX Live, MiKTeX などの TeX ディストリビューション)
    * `Synapsen Ersteller` でのPDFビルドに必須です。
    * 導入方法は、こちらの解説記事などを参考にしてください。<br>
        → **[LaTeXの環境構築 \~VSCodeでLaTeXを使いたいだけなのに TeX Liveの導入が必要なのは何故?\~](https://qiita.com/Kurato-Tsukishiro/items/58232e619a1878692bed)**
* **(オプション) Tesseract OCR**
    * `Normalisierer` で画像PDFのOCR（本文テキスト抽出）機能 を使う場合に必要です。
    * インストール後、`pytesseract` が認識できるようPATHを通してください。
    * 導入方法は、こちらの解説記事などを参考にしてください。<br>
        → **[画像から文字を瞬時に読み取る！Tesseractとpytesseractの驚異の力【Python】](https://qiita.com/ryome/items/16fc42854fe93de78a2f)**
* **(オプション) Pandoc**
    * `Normalisierer` でMarkdown (.md) ファイルのPDF変換機能を使う場合に必要です。
    * `Install.bat` により、`winget` を使用して自動インストールを試みます。
    * 導入方法は、こちらの公式ページなどを参考にしてください。<br>
        → **[Pandoc - Installing](https://pandoc.org/installing.html)**
* **Pythonライブラリ**: ( `requirements.txt` 参照)
    * [**customtkinter**](https://github.com/TomSchimansky/CustomTkinter) (MIT License) - GUI構築用
    * [**pandas**](https://github.com/pandas-dev/pandas) (BSD-3-Clause License) - 索引データの管理・検索用
    * [**PyMuPDF (fitz)**](https://github.com/pymupdf/PyMuPDF) (AGPL-3.0 License) - PDFの正規化・情報抽出用 (※プロジェクト全体のAGPLライセンスの要因)
    * [**pypdf**](https://github.com/py-pdf/pypdf) (BSD-3-Clause License) - PDFの統合・正規化用
    * [**Pillow**](https://github.com/python-pillow/Pillow) (HPND License) - OCR処理のための画像操作用
    * [**pytesseract**](https://github.com/madmaze/pytesseract) (Apache-2.0 License) - Tesseract OCRエンジン連携用
    * [**networkx**](https://github.com/networkx/networkx) (BSD-3-Clause License) - グラフ・ビジュアライゼーションのデータ構築用
    * [**pyvis**](https://github.com/WestHealth/pyvis) (MIT License) - インタラクティブなグラフ描画用
    * [**tkinterdnd2**](https://github.com/Eliav2/tkinterdnd2) (MIT License) - `Normalisierer` でD&D機能を実現するため
    * [**playwright**](https://github.com/microsoft/playwright-python) (Apache-2.0 License) - `Normalisierer` でWebクリップとMarkdown変換を実現するため

## セットアップ

1.  **Synapsenのダウンロード:**
    * リポジトリの [Releasesページ](https://github.com/Kurato-Tsukishiro/Synapsen/releases) から、最新のリリース（`.zip` (ソースコード 又は `.exe` が含まれるパッケージ)）をダウンロードします。

2.  **PDFテンプレートの入手 (推奨):**
    * `Synapsen` をより便利に使うため、**専用のPDFテンプレート**（`DotLegalPad_Template-A4_Form.pdf` など）の使用を推奨します。
    * これを使うと、ノート作成時にIndex Keyを選ぶだけで、後で `Ersteller` が自動で読み取ってくれるため、**手動でIndex Keyを登録する手間が省けます**。
    * [Releasesページ](https://github.com/Kurato-Tsukishiro/Synapsen/releases) から、以下のPDFテンプレートファイル（`.pdf`）をダウンロードしてください。
        1.  **フォーム付き (`..._Form.pdf`):** Index Keyを選択するプルダウンが付いたPDF。QUADERNOでは「ドキュメント」として扱われ、**ページ追加ができません**。
        2.  **フォーム無し (`...Template.pdf`):** ページ追加が可能な、通常の「ノート」テンプレート。
    * ※ これらのテンプレートは `CC0 (パブリックドメイン)` です。自由にコピー、改変、再配布して構いません。

3.  **ライブラリのインストール:**
    * ダウンロードしたフォルダにある `Install.bat` をダブルクリックして実行し、必要なPythonライブラリ、Webクリップ用ブラウザ、およびMarkdown変換用のPandocをインストールします。
    * （または、コマンドプロンプトで `pip install -r requirements.txt` と `playwright install chromium` を実行し、`Pandoc`を手動でインストールします）

4.  **`config.ini` の設定:**
    * フォルダ内にある `config.ini` を開き、**最低限 `[Paths]` セクションのパス**を、ご自身の環境に合わせて編集します。
    * ※ 推奨テンプレート (`DotLegalPad`) を使用する場合、`[Extraction]` や `[CommonplaceKeys]` は、リリースに同梱されている `config.ini` のデフォルト設定から**変更不要**です。
    * ※ **テンプレートを使わない場合**は、ご自身で `[Extraction]` の座標を調べるか、`Ersteller` でノートごとに手動でIndex Keyを登録する必要があります。

    **`config.ini` の設定項目 (空のテンプレート):**
    ```ini
    [Paths] # 絶対パス 又は config.ini からの相対パスを指定
    # 事前定義タグを保存しているテキストファイルのパス
    tags_data_path = 
    
    # Normaliiererが(フォームのテキスト化で)使用するフォントファイルのフルパス
    # Noto San JP を使用する場合は "%LOCALAPPDATA%\Microsoft\Windows\Fonts\NotoSansJP-Regular.otf" を使用して下さい
    font_path = 
    
    # Nexusでの情報表示に使用するマスターDBのパス
    database_path = 
    
    # マスターDBが存在するフォルダ下に統合PDFが存在しない場合に NexusがPDFを開く為に検索するフォルダのパス
    pdf_root_folder = 
    
    [Automation]
    # Synapse Ersteller で統合PDFを生成した際、
    # [Paths]のdatabase_pathで指定されたマスターDBに、目次情報を自動で「追記」するか (true/false) (主にDebug用設定)
    auto_append_to_default_db = 
    
    # 上記有効時、目次情報を個別で「保存」するか (true/false)
    create_individual_csv = 
    
    # Normalisierer で Tesseract OCR (低速な光学文字認識) を実行するか (true/false)
    # false の場合でも、PDFに埋め込まれた既存のテキスト抽出（高速）は実行されます。
    # Tesseract-OCR をPCにインストールしていない場合は false にしてください。
    enable_tesseract_ocr = 
    
    [LaTeX]
    # 正規化及び統合の用紙サイズの指定 (A4/A5)
    paper_size = 
    
    # PDF生成時に使用するフォント名
    # font = Noto Sans JP
    font = 
    
    # PDFのプロパティに表示される著者名
    author = 
    
    # PDFのタイトル接頭辞（この後ろに「(YYYY年 M月)」が付きます）
    title_prefix = 
    
    [Extraction]
    # Erstrller で読み取り Index Keyを取得する範囲 (DotLegalPadテンプレートの座標)
    key_rect = 
    
    [CommonplaceKeys]
    # Index Key の設定 (DotLegalPadテンプレートの選択肢はこれと一致させる)
    options = 
    
    [KeyIcons]
    # = の左側にキー、右側に表示したいアイコン（Unicode絵文字など）を記述
    
    [KeyColors]
    # = の左側にキー、右側に表示したい色（16進数カラーコード）を記述
    
    [Search]
    # オートコンプリートの候補に、全ノートで使用されているタグを含めるか (true/false)
    # true: 全ノートのタグ + 事前定義タグ (「野良タグ」も再利用可能になります)
    # false: 事前定義タグ(PDFTags.txt)のみ (ロードが高速で、候補が整理されます)
    include_all_tags_for_autocomplete = 

    ```

    <details>
    <summary><b>▼ クリックして推奨設定例 (`config.ini` のデフォルト値) を表示</b></summary>
    
    ```ini
    [Paths] 
    tags_data_path = PDFTags.txt
    font_path = C:\windows\fonts\msgothic.ttc
    database_path = Synapsen_Master.db
    pdf_root_folder = ./
    
    [Automation]
    auto_append_to_default_db = true
    create_individual_csv = false
    enable_tesseract_ocr = false
    
    [LaTeX]
    paper_size = A4
    font = MS UI Gothic
    author = Synapsen Ersteller
    title_prefix = 月刊 統合ノート
    
    [Extraction]
    # Erstrller で読み取り Index Keyを取得する範囲
    key_rect = 0, 13, 391, 73
    
    [CommonplaceKeys]
    # Index Key の設定
    options = タスク,アイデア,思考・考察,コミュニケーション,学習・情報収集,日常・その他
    
    [KeyIcons]
    # = の左側にキー、右側に表示したいアイコン（Unicode絵文字など）を記述
    タスク = ♥
    アイデア = ♥
    思考・考察 = ♥
    コミュニケーション = ♥
    学習・情報収集 = ♥
    日常・その他 = ♥
    
    [KeyColors]
    # = の左側にキー、右側に表示したい色（16進数カラーコード）を記述
    # アプリ内のリスト表示・統合ノートのヘッダーおよび索引で使用されます
    タスク = #FE0000
    アイデア = #FFFF02
    思考・考察 = #8802FF
    コミュニケーション = #02FF01
    学習・情報収集 = #02FFFF
    日常・その他 = #F2F2F2
    
    [Search]
    include_all_tags_for_autocomplete = true

    ```
    </details>

## 使い方

`Synapsen` はPDFテンプレートを使わなくても利用できます。その場合、`Ersteller` で「フォルダから新規読み込み」を行った後、リストから各ノートをクリックして、手動で「Index Key」を割り当ててください。

以下は、推奨テンプレート（`DotLegalPad` など）を使用して、Index Keyの入力を自動化する推奨ワークフローです。

---

### ステップ0: ノートの作成 (テンプレート利用＠QUADERNO)

QUADERNOでは「フォーム付きPDF（ドキュメント）」にページを追加できません。以下の手順で2種類のテンプレートを使い分けます。

1.  **1ページ目 (Index Keyの指定):**
    * **フォーム付きPDF** (`..._Form.pdf`) をQUADERNO上で**複製**して、新しいノート（ドキュメント）とします。
    * ファイル名を `YYYYMMDD_hhmmss_タイトル.pdf` の形式に変更します。
    * ノート左上のプルダウンメニューから、そのノートの「Index Key」（例: 'アイデア'）を選択し、1ページ目の内容を書き込みます。

2.  **2ページ目以降 (ページの追加):**
    * **フォーム無しPDF** (`...Template.pdf`) を使って、QUADERNOの「**サイドノートを作成**」機能でページを追加します。
    * `Synapsen Ersteller` は、このサイドノート (`..._Note.pdf`) を自動で親ノートと紐付け、Index Keyを継承させます。

3.  **PCへのエクスポート:**
    * 書き終わったら、1ページ目のドキュメント（`..._Form.pdf` を複製したもの）と、2ページ目以降のサイドノート（`..._Note.pdf`）の両方をPCにエクスポートします。
    * (ScanSnapユーザーは、スキャンしたPDFを直接エクスポートフォルダに保存してください)

### ステップ1: 正規化 (Normalisierer)

1.  `Synapsen_Normalisierer_main.py` を実行します。
2.  以下のいずれかの方法でファイルを処理します。
    * **A) フォルダ一括処理:**
        * 「入力/出力フォルダを選んで処理実行」をクリックします。
        * **入力元フォルダ**（スキャン及びクアデルノで作成したPDF、または .md ファイルがある場所）を選択します。
        * **出力先フォルダ**（正規化済みPDFを保存する場所）を選択します。
    * **B) D&D / ペースト:**
        * 「D&D / ペースト」ボタンを押し、別ウィンドウを開きます。
        * IndexKeyやコメントを入力し、ファイル（PDF/JPG/PNG/MD）をD&Dするか、スクリーンショットをCtrl+Vで貼り付けます。
        * 「出力先を選んで処理実行」をクリックします。
    * **C) Webクリップ:**
        * 「Webクリップ」ボタンを押し、別ウィンドウを開きます。
        * URLを入力し「ページ情報取得」をクリックします。（※PDF/画像への直接リンクにも対応）
        * IndexKey、コメント、書誌情報を入力・編集します。
        * 「出力先を選んでクリップ実行」をクリックします。<br><br>
    * **補遺:**
        > 「D&D/ペースト」および「Webクリップ」機能は、実行時に「入力/出力フォルダを選んで処理実行」機能（上記A）と**同一の正規化処理（フォームのテキスト化、サイズ統一、OCR処理など）を自動的に実行します。**
        >
        > そのため、これらの機能で出力されたPDFを `Normalisierer` で再度処理する必要はなく、**そのまま `Ersteller` に読み込ませて使用できます。**

### ステップ2: 統合 (Ersteller)
1.  `Synapsen_Ersteller_main.py` を起動します。
2.  正規化済みのフォルダを読み込み、メタデータを編集します。
3.  「統合PDFを生成」で月ごとのPDFを作成し、データベースに登録します。

* 補遺:
  * メタデータの自動読み込みについて
    * ファイル名 (`YYYYMMDD_hhmmss_...`) から日付とタイトルが読み込まれます。
    * PDF内容 (`key_rect` の座標) から「Index Key」が自動で読み込まれます。
      * サイドノート (`..._Note.pdf`) にも親のIndex Keyが自動で継承されます。
  * `リスト保存 (CSV)` / `リスト読込 (CSV)` について
    * 現在の編集状態（メタデータのみ）を中間ファイルとして保存/読み込みできます。
    * この中間ファイルには、本文テキストは含まれません
      * 本文テキストの抽出は「統合PDFを生成」を実行した時のみに行われます。
      * 本文テキストが保存されるファイルは、**マスターDB**のみです。

### ステップ3: 閲覧 (Nexus)
1.  `Synapsen_Nexus_main.py` を起動します。
2.  検索やフィルターでノートを探し、詳細やプレビューを確認します。
3.  「グラフ表示」や「関連グラフ」で知識のつながりを視覚化します。
4.  必要なノートを複数選択し、「リンクコピー」でMOCを作成したり、「エクスポート」でPDFとして書き出したりして活用します。

* 補遺:
  * アプリは `config.ini` で指定された**マスターDB**を自動で読み込みます。
  * 検索文字列の例 :
    * `fulltext:Python`
    * `ikey:タスク AND (アイデア OR 思考)`
    * `(date:>=20240101 AND date:<=20251231) AND (ikey:アイデア OR ikey:学習・情報収集)`
  * ``tag: ``プレフィクスを使用した時、ノートに登録されているタグからの予測変換が呼び出されます。

---

## PDFテンプレートのカスタマイズ (上級者向け)

配布されているPDFテンプレート (`DotLegalPad_Template-A4_Form.pdf` など) は、リポジトリ内のPythonスクリプトによって生成されています。
Index Keyの選択肢（タグ）、フォント、レイアウトなどを変更したい場合は、以下の手順で自分専用のテンプレートを生成できます。

1.  **設定ファイルの編集:**
    * `PDF_Templates/DotLegalPad/DotLegalPad_Config.py` をテキストエディタで開きます。
    * `OPTIONS = [...]` のリストを編集すると、PDFのプルダウンメニューに表示されるIndex Keyの選択肢を変更できます。
    * `FONT_PATH` や `COLOR_LINE` などの他の設定も、好みに合わせて変更できます。

2.  **ライブラリの確認:**
    * テンプレートの生成には `PyMuPDF` ライブラリ が必要です。
    * セットアップ時に `Install.bat` を実行済みであれば、必要なライブラリはすでにインストールされています。

3.  **生成スクリプトの実行:**
    * コマンドプロンプト（ターミナル）で `PDF_Templates/DotLegalPad/` フォルダに移動します。
    * `python Generate_DotLegalPad_Form_Template.py` を実行すると、フォーム付きのPDFが生成されます。
    * `python Generate_DotLegalPad.py` を実行すると、フォーム無しの（ページ追加用）PDFが生成されます。

> **重要:**
> `DotLegalPad_Config.py` の `OPTIONS` を変更した場合、Synapsen本体の `config.ini` の `[CommonplaceKeys]` セクションにある `options = ...` の内容も、必ず一致させてください。

> **ライセンスに関する注意:**
> このセクションで変更する `.py` ファイル（生成スクリプト）は、`AGPL-3.0` ライセンスの対象です。
> もし、あなたが**変更を加えた生成スクリプト自体を再配布・公開する場合**は、`AGPL-3.0` の条項に従う必要があります（生成された `.pdf` ファイルは `CC0` のため、自由に配布して問題ありません）。

---

## ライセンス

### ソースコード (AGPL-3.0)
このプロジェクトの**ソースコード**（(`PDF_Templates` 内の生成スクリプトも含む)`.py` ファイル）は、**GNU Affero General Public License v3.0 (AGPL-3.0)** の下でライセンスされています。<br><br>
これは、`Synapsen` の中核機能において、AGPL-3.0 ライセンスである `PyMuPDF (fitz)` ライブラリ を使用しているためです。<br>
AGPL-3.0の条項に基づき、このライブラリを利用する本アプリケーション全体も同じライセンスに従います。<br><br>
詳細は、同梱されている `LICENSE` ファイルを参照してください。

---

### アイコンおよびグラフィックアセット (CC BY-SA 4.0)
このリポジトリの **`assets/` フォルダ** に含まれるすべてのファイル（ロゴ、アイコン、`.png` 画像、および `.gvdesign` ソースファイル）は、ソースコードとは別に **Creative Commons Attribution-ShareAlike 4.0 International (CC BY-SA 4.0)** の下でライセンスされています。<br><br>
詳細は、同梱されている `LICENSE-ASSETS.md` ファイルを参照してください。

---

### PDFテンプレート (CC0 - パブリックドメイン)
[Releasesページ](https://github.com/Kurato-Tsukishiro/Synapsen/releases) や `PDF_Templates/PDF/` フォルダで配布されている **`.pdf` テンプレートファイル**（`DotLegalPad_Template-A4_Form.pdf` など）は、**CC0 (パブリックドメイン)** です。<br>
これらのテンプレート（およびそれに書き込んだあなたのノート）は、ライセンスを一切気にすることなく、自由にコピー、改変、共有、再配布が可能です。

---

### 生成されるグラフ (synapsen_graph.html) について
`Synapsen Nexus` が「グラフ表示」機能 で生成する `synapsen_graph.html` 及び エクスポート機能 で生成する `relation_graph.html` ファイルは、`Synapsen` プログラムの「出力」であり、AGPL-3.0 ライセンス の対象外です。<br><br>
このHTMLファイルは、`pyvis` ライブラリ（MITライセンス） と、ユーザー自身のノートデータ（タイトルやリンク構造）で構成されています。<br>
したがって、ユーザーはこれを ``AGPL-3.0`` を気にすることなく自由に利用・公開・配布できます。

---

## 謝辞 (Acknowledgements)
このソフトウェアは、多くの優れたオープンソースライブラリによって実現しています。<br><br>
特に、GUI構築のための **CustomTkinter**、<br>
データ操作のための **pandas**、<br>
PDF処理の中核を担う **PyMuPDF** と **pypdf**、<br>
OCR機能を実現する **Pillow** と **pytesseract**、<br>
知識グラフの可視化を実現する **NetworkX** と **Pyvis**、<br>
D&D機能を実現する **tkinterdnd2**、<br>
そしてWebクリップ機能とMarkdown変換を実現する **Playwright** および **Pandoc** の開発者コミュニティに心から感謝申し上げます。<br><br>
また、このプロジェクトの設計、コード作成、リファクタリング、およびドキュメント整備は、GoogleのAIである **Gemini** の支援を受けて行われました。
