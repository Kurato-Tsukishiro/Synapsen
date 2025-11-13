import networkx as nx
from pyvis.network import Network
import re
import sys
import webbrowser
from pathlib import Path
from textwrap import dedent
from utils import get_pdf_uri_for_note


class GraphManager:
    @staticmethod
    def generate_graph_html(
        df,
        key_icons,
        key_colors,
        loaded_db_path,
        pdf_root_folder,
        output_path=None
    ):
        """
        DataFrameからネットワークグラフを生成し、HTMLファイルとして保存する。

        Returns:
            Path: 保存されたHTMLファイルのパス。
        """
        # 1. グラフ構築 (NetworkX)
        G = nx.DiGraph()
        notes_in_graph = set(df['key'])
        link_pattern = re.compile(r"\[\[(.*?)\]\]")

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
        for index, row in df.iterrows():
            source_key = row.get('key')
            memo = row.get('memo', '')

            for match in link_pattern.finditer(memo):
                full_match_content = match.group(1).strip()
                target_key = full_match_content.split(':')[0].strip()

                if target_key in notes_in_graph:
                    if source_key != target_key:
                        G.add_edge(source_key, target_key)
                        edge_count += 1

        print(f"[GraphManager] グラフ生成: {len(df)} ノード, {edge_count} エッジ")

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
            print(f"[GraphManager] JS注入エラー: {e}")

    @staticmethod
    def open_graph(html_path):
        """指定されたHTMLをブラウザで開く"""
        try:
            webbrowser.open(Path(html_path).as_uri())
        except Exception as e:
            print(f"[GraphManager] ブラウザ起動エラー: {e}")
