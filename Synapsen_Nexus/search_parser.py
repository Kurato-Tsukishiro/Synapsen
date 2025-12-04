import re

import logging

logger = logging.getLogger(__name__)


def split_respecting_parens(query, operator):
    """
    括弧を考慮して、トップレベルの演算子でのみ分割する。

    (例: 'A AND (B OR C)' を ' AND ' で分割すると ['A', '(B OR C)'] となる)

    Args:
        query (str): 分割対象の検索クエリ。
        operator (str): 分割に使用する演算子 (' AND ' または ' OR ')。

    Returns:
        list[str]: 分割されたクエリ文字列のリスト。
    """
    parts = []
    balance = 0
    current_part = ""
    op_len = len(operator)

    i = 0
    while i < len(query):
        char = query[i]
        if char == "(":
            balance += 1
            current_part += char
            i += 1
        elif char == ")":
            balance -= 1
            current_part += char
            i += 1
        # 括弧の外 (balance == 0) で、大文字小文字を無視して演算子をチェック
        elif balance == 0 and query[i : i + op_len].upper() == operator.upper():
            parts.append(current_part.strip())
            current_part = ""
            i += op_len
        else:
            current_part += char
            i += 1

    parts.append(current_part.strip())
    return [p for p in parts if p]


def _build_term_sql(term, include_full_text=False):
    """
    単一の検索語を SQL 条件式とパラメータに変換する。

    Args:
        df (pd.DataFrame): 検索対象のDataFrame。
        term (str): 単純な検索語 (例: 'Python', 'title:Python', 'date:>=20240101')。
        include_full_text (bool, optional): 本文及びメモ検索が有効かどうか。デフォルトは False。

    Returns: (sql_condition, params_list)
    """
    term = term.strip()

    # リンク形式 [[Key]] の完全一致
    link_match = re.match(r"^\[\[(\d{14,})(?::.*)?\]\]$", term)
    if link_match:
        return "key = ?", [link_match.group(1)]

    # プレフィックスの解析
    target_column = None
    search_value = term

    # 検索対象カラムのマッピング
    col_map = {
        "title": "title",
        "key": "key",
        "date": "date",
        "time": "time",
        "tag": "tags",
        "tags": "tags",
        "memo": "memo",
        "cpkey": "commonplace_key",
        "indexkey": "commonplace_key",
        "ikey": "commonplace_key",
        "fulltext": "full_text",
        "text": "full_text",
        "filename": "merged_pdf_filename",
        "file": "merged_pdf_filename",
    }

    # プレフィックス (key:など) があるかチェック
    if ":" in term:
        parts = term.split(":", 1)
        prefix = parts[0].lower().strip()
        val = parts[1].strip()

        if prefix in col_map and val:
            target_column = col_map[prefix]
            search_value = val

    if not search_value:
        return "1=1", []  # 無効な条件は無視

    # --- カラム別の SQL 生成 ---

    # 1. Date (範囲検索対応)
    if target_column == "date":
        # YYYYMMDD-YYYYMMDD
        range_match = re.match(r"^(\d{8})-(\d{8})$", search_value)
        if range_match:
            return "date BETWEEN ? AND ?", [range_match.group(1), range_match.group(2)]

        # >=YYYYMMDD
        gte_match = re.match(r"^>=(\d{8})$", search_value)
        if gte_match:
            return "date >= ?", [gte_match.group(1)]

        # <=YYYYMMDD
        lte_match = re.match(r"^<=(\d{8})$", search_value)
        if lte_match:
            return "date <= ?", [lte_match.group(1)]

        # 部分一致 (LIKE)
        return "date LIKE ?", [f"%{search_value}%"]

    # 2. Time (ワイルドカード検索)
    elif target_column == "time":
        # ユーザー入力の '_' を SQL のワイルドカード '_' に、不足分を補完
        # SQLite の LIKE でも '_' は「任意の1文字」なのでそのまま使える
        pattern = search_value
        if len(pattern) < 6:
            pattern = pattern.ljust(6, "_")
        elif len(pattern) > 6:
            pattern = pattern[:6]
        return "time LIKE ?", [pattern]

    # 3. FTS (全文検索)
    elif target_column == "full_text":
        # notes_fts テーブルを活用 (rowid = key)
        return "key IN (SELECT rowid FROM notes_fts WHERE notes_fts MATCH ?)", [
            search_value
        ]

    # 4. 特定カラムの部分一致
    elif target_column:
        return f"{target_column} LIKE ?", [f"%{search_value}%"]

    # 5. グローバル検索 (複数カラム OR)
    else:
        # FTS が有効なら本文も含める
        conditions = [
            "title LIKE ?",
            "tags LIKE ?",
            "key LIKE ?",
            "time LIKE ?",
            "commonplace_key LIKE ?",
            "date LIKE ?",
        ]
        params = [f"%{search_value}%"] * 6

        if include_full_text:
            conditions.append(
                "key IN (SELECT rowid FROM notes_fts WHERE notes_fts MATCH ?)"
            )
            params.append(search_value)

        return f"({' OR '.join(conditions)})", params


def parse_query_to_sql(query, include_full_text=False):
    """
    クエリ文字列を解析し、SQLite用の WHERE 句とパラメータを生成する。

    Returns:
        tuple: (where_clause_string, parameters_list)
        例: ("(tags LIKE ?) AND (date >= ?)", ['%python%', '20240101'])
    """
    if not query or not query.strip():
        return "", []

    # OR で分割
    or_parts = split_respecting_parens(query, " OR ")
    or_sql_parts = []
    or_params = []

    for or_part in or_parts:
        # AND で分割
        and_parts = split_respecting_parens(or_part, " AND ")
        and_sql_parts = []
        and_params = []

        for term in and_parts:
            term = term.strip()
            # 括弧の処理 (再帰)
            if term.startswith("(") and term.endswith(")"):
                inner_sql, inner_params = parse_query_to_sql(
                    term[1:-1], include_full_text
                )
                if inner_sql:
                    and_sql_parts.append(f"({inner_sql})")
                    and_params.extend(inner_params)

            # NOT検索 (-)
            elif term.startswith("-"):
                term_body = term[1:].strip()
                # 孤立ノートのNOTは「孤立していない」
                if term_body.lower() == "is:orphan":
                    and_sql_parts.append(
                        "key IN (SELECT source_key FROM note_links UNION"
                        "SELECT target_key FROM note_links)"
                    )
                else:
                    s, p = _build_term_sql(term_body, include_full_text)
                    if s:
                        and_sql_parts.append(f"NOT ({s})")
                        and_params.extend(p)

            # 孤立ノート (is:orphan)
            elif term.lower() == "is:orphan":
                and_sql_parts.append(
                    "key NOT IN (SELECT source_key FROM note_links UNION"
                    "SELECT target_key FROM note_links)"
                )

            # 通常検索
            else:
                s, p = _build_term_sql(term, include_full_text)
                if s:
                    and_sql_parts.append(s)
                    and_params.extend(p)

        if and_sql_parts:
            or_sql_parts.append(f"({' AND '.join(and_sql_parts)})")
            or_params.extend(and_params)

    if not or_sql_parts:
        return "", []

    return f"({' OR '.join(or_sql_parts)})", or_params
