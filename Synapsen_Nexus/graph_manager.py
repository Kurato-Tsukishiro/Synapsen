import networkx as nx
from pyvis.network import Network
import sys
import webbrowser
from pathlib import Path
from textwrap import dedent
from utils import get_pdf_uri_for_note
import sqlite3

import logging
logger = logging.getLogger(__name__)


class GraphManager:
    @staticmethod
    def generate_graph_html(
        df,
        key_icons,
        key_colors,
        loaded_db_path,
        pdf_root_folder,
        output_path=None,
        db_conn=None  # NexusからDB接続を受け取る
    ):
        """
        DataFrameからネットワークグラフを生成し、HTMLファイルとして保存する。

        Returns:
            Path: 保存されたHTMLファイルのパス。
        """
        # グラフ構築 (NetworkX)
        G = nx.DiGraph()
        keys_in_graph = set(df['key'])

        # グラフ化対象のリンク(edge)をDBから一括取得
        edges_data = []
        if (db_conn or loaded_db_path) and keys_in_graph:
            conn_to_use = None
            created_conn = False
            try:
                if db_conn:
                    # Nexusから渡された読み取り専用接続を使用
                    conn_to_use = db_conn
                else:
                    # (ExportManagerなどから呼ばれた場合) 自分で接続を作成
                    conn_to_use = sqlite3.connect(
                        f"file:{loaded_db_path}?mode=ro", uri=True)
                    created_conn = True

                cursor = conn_to_use.cursor()

                # プレースホルダ (?) をキーの数だけ生成
                placeholders = ','.join('?' for _ in keys_in_graph)

                # リンク元(source)がグラフ対象に含まれるリンクのみ取得
                sql = (
                    f"SELECT source_key, target_key FROM note_links WHERE "
                    f"source_key IN ({placeholders})"
                )

                cursor.execute(sql, tuple(keys_in_graph))
                edges_data = cursor.fetchall()

                logger.info(
                    f"[GraphManager] DBから {len(edges_data)} 件のリンク(エッジ)を取得しました。"
                )

            except Exception as e:
                logger.error(f"[GraphManager] DBからのリンク取得に失敗: {e}")
            finally:
                # 自分で作成した接続のみ閉じる
                if created_conn and conn_to_use:
                    conn_to_use.close()

        # ノードを追加
        for index, row in df.iterrows():
            key = row.get('key')
            title = row.get('title', 'N/A')
            cp_key = row.get('commonplace_key', '').lower()

            icon_code = key_icons.get(cp_key, '•')
            color_hex = key_colors.get(cp_key, '#FFFFFF')

            # PDFへのURIを取得
            file_uri = get_pdf_uri_for_note(
                row, loaded_db_path, pdf_root_folder
            )

            tooltip = f"Key: {key}\nIndex: {cp_key}"
            if file_uri:
                tooltip += "\n(ダブルクリックしてPDFを開く)\n(右クリックでKeyをコピー)"

            G.add_node(
                key,
                label=title,
                title=tooltip,
                shape='icon',
                icon={
                    'code': icon_code,
                    'color': color_hex,
                    'size': 40
                },
                color=color_hex,
                pdf_url=file_uri if file_uri else ""
            )

        # エッジを追加
        edge_count = 0

        # DBから取得した edges_data をイテレート
        for source_key, target_key in edges_data:
            # リンク先(target)もグラフ描画対象(keys_in_graph)に含まれるか確認
            if target_key in keys_in_graph:
                if source_key != target_key:
                    G.add_edge(source_key, target_key)
                    edge_count += 1

        logger.info(f"[GraphManager] グラフ生成: {len(df)} ノード, {edge_count} エッジ")

        # 2. 視覚化設定 (Pyvis)
        nt = Network(
            height="95vh",
            width="100%",
            bgcolor="#222222",
            font_color="white",
            directed=True,
            notebook=False
        )
        nt.from_nx(G)

        # 物理演算設定
        nt.set_options("""
        var options = {
          "physics": {
            "solver": "barnesHut",
            "barnesHut": {
              "gravitationalConstant": -8000,
              "centralGravity": 0.3,
              "springLength": 95,
              "springConstant": 0.04,
              "damping": 0.09
            },
            "minVelocity": 0.75
          },
          "interaction": {
            "tooltipDelay": 200,
            "hideEdgesOnDrag": true,
            "hover": true,
            "hoverConnectedEdges": true,
            "selectConnectedEdges": true,
            "navigationButtons": true,
            "keyboard": { "enabled": true }
          },
          "edges": {
            "arrows": {
              "to": { "enabled": true, "scaleFactor": 0.5 }
            },
            "color": {
              "color": "#848484",
              "highlight": "#FFFFFF",
              "hover": "#DDDDDD",
              "inherit": false
            },
            "smooth": {
              "type": "continuous",
              "forceDirection": "none",
              "roundness": 0.5
            }
          }
        }
        """)

        # 3. 保存パスの決定
        if output_path:
            save_path = Path(output_path)
        else:
            # デフォルトパス (アプリ実行位置基準)
            if getattr(sys, 'frozen', False):
                base_path = Path(sys.executable).parent
            else:
                base_path = Path(__file__).parent.parent
            save_path = base_path / "synapsen_graph.html"

        # 4. HTML書き出しとJS注入
        nt.save_graph(str(save_path))
        GraphManager._inject_custom_js(save_path)

        return save_path

    @staticmethod
    def _inject_custom_js(html_path):
        """生成されたHTMLにカスタムインタラクション用JSを注入する"""
        custom_js = dedent("""
        document.addEventListener('DOMContentLoaded', function() {
            if (typeof network !== 'undefined') {
                // ダブルクリック: PDFを開く
                network.on("doubleClick", function(properties) {
                    var { nodes } = properties;
                    if (nodes.length > 0) {
                        var nodeId = nodes[0];
                        var nodeData = this.body.nodes[nodeId].options;
                        if (nodeData.pdf_url && nodeData.pdf_url !== "") {
                            window.open(nodeData.pdf_url, '_blank');
                        }
                    }
                });
                // 右クリック: Keyをコピー
                network.on("oncontext", function (params) {
                    params.event.preventDefault();
                    var nodeId = network.getNodeAt(params.pointer.DOM);
                    if (nodeId) {
                        var textToCopy = nodeId;
                        // クリップボードAPI (HTTPS/localhost または特定ブラウザ設定が必要)
                        if (navigator.clipboard && window.isSecureContext) {
                            navigator.clipboard.writeText(textToCopy).then(
                                function() {
                                    // 成功時のフィードバック
                                    alert('Keyをコピーしました: ' + textToCopy);
                                },
                                function(err) {
                                    // 失敗時 -> プロンプトを表示
                                    prompt(
                                        "コピーしてください (Ctrl+C):",
                                        textToCopy
                                    );
                                }
                            );
                        } else {
                            prompt("コピーしてください (Ctrl+C):", textToCopy);
                        }
                    }
                });
            }
        });
        """)

        try:
            with open(html_path, 'r', encoding='utf-8') as f:
                content = f.read()

            script_tag = (
                "<script type=\"text/javascript\">\n" +
                f"{custom_js}\n</script>\n</head>"
            )
            content = content.replace("</head>", script_tag, 1)

            with open(html_path, 'w', encoding='utf-8') as f:
                f.write(content)
        except Exception as e:
            logger.error(f"[GraphManager] JS注入エラー: {e}")

    @staticmethod
    def open_graph(html_path):
        """指定されたHTMLをブラウザで開く"""
        try:
            webbrowser.open(Path(html_path).as_uri())
        except Exception as e:
            logger.error(f"[GraphManager] ブラウザ起動エラー: {e}")
