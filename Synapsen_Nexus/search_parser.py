import pandas as pd
import re

import logging
logger = logging.getLogger(__name__)  # <--- ロガーを取得


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
        if char == '(':
            balance += 1
            current_part += char
            i += 1
        elif char == ')':
            balance -= 1
            current_part += char
            i += 1
        # 括弧の外 (balance == 0) で、大文字小文字を無視して演算子をチェック
        elif balance == 0 and query[i:i+op_len].upper() == operator.upper():
            parts.append(current_part.strip())
            current_part = ""
            i += op_len
        else:
            current_part += char
            i += 1

    parts.append(current_part.strip())
    return [p for p in parts if p]  # 空の文字列を除外


def evaluate_simple_term(df, term, include_full_text=False, db_conn=None):
    """
    プレフィックス検索、またはグローバル検索を実行する。

    'title:Python' のようなプレフィックス検索、または 'Python' のような
    グローバル検索を処理し、該当する行のboolマスク (pd.Series) を返す。
    ['date'] プレフィックスの日時範囲検索 (>=, <=, YYYYMMDD-YYYYMMDD) にも対応。

    'full_text' または 'memo' を対象とする検索は、db_conn を使用して
    FTS MATCH ではなく 'LIKE' でDBを直接検索する。

    Args:
        df (pd.DataFrame): 検索対象のDataFrame。
        term (str): 単純な検索語 (例: 'Python', 'title:Python', 'date:>=20240101')。
        include_full_text (bool, optional): 本文及びメモ検索が有効かどうか。デフォルトは False。
        db_conn (sqlite3.Connection, optional): SQLiteのデータベース接続オブジェクト。
                                                デフォルトは None。

    Returns:
        pd.Series: クエリに一致した行がTrueとなるboolマスク。
    """
    # [[Key]] 形式のリンクを検出
    # (例: [[20241117000000]] や [[20241117000000: Title]] )
    # この形式は「完全一致」検索として扱う
    link_pattern = re.match(r'^\[\[(\d{14,})(?::.*)?\]\]$', term.strip())
    if link_pattern:
        extracted_key = link_pattern.group(1)
        # Key検索は「完全一致」
        return df['key'] == extracted_key

    search_fields_map = {
            'title': 'title',
            'key': 'key',
            'date': 'date',
            'time': 'time',
            'tag': 'tags',
            'tags': 'tags',             # 'tag'でも'tags'でも検索可
            'memo': 'memo',
            'cpkey': 'commonplace_key',
            'indexkey': 'commonplace_key',
            'ikey': 'commonplace_key',  # IndexKeyのエイリアス
            'fulltext': 'full_text',
            'text': 'full_text'         # fulltextのエイリアス
    }

    target_column = None
    search_value = term

    # プレフィックス (key:など) があるかチェック
    if ':' in term:
        parts = term.split(':', 1)
        prefix = parts[0].lower().strip()
        value = parts[1].strip()

        if prefix in search_fields_map and value:
            target_column = search_fields_map[prefix]
            search_value = value

    if not search_value:
        # 検索語が空なら、何もヒットしないマスクを返す
        return pd.Series([False] * len(df), index=df.index)

    term_condition = pd.Series([False] * len(df), index=df.index)

    # --- 【FTS検索判定 (LIKEフォールバック)】 ---
    # 検索語がDB検索(LIKE)の対象か？
    # 1. 'fulltext:', 'text:', 'memo:' プレフィックス ('本文・メモ検索'の状態によらない)
    is_db_column = (target_column in ('full_text', 'memo'))

    # 2. プレフィックスがなく (global) + '本文・メモ検索' ON
    is_db_global = (target_column is None and include_full_text)

    if (is_db_column or is_db_global) and db_conn:
        # DBを 'LIKE' で検索 (低速だが確実)
        try:
            # LIKE検索用の検索語 ( %query% )
            like_query_term = f'%{search_value}%'

            logger.warning(
                f"[FTS DEBUG] DB 'LIKE'検索を実行中... (Term: {like_query_term}, "
                f"is_db_column: {is_db_column}, is_db_global: {is_db_global})",
                extra={'sensitive': True}
            )

            # どの列を検索するか
            if target_column == 'full_text':
                # 'fulltext:query' -> notes.full_text 列のみ検索
                sql = "SELECT key FROM notes WHERE full_text LIKE ?"
                params = (like_query_term,)

            elif target_column == 'memo':
                # 'memo:query' -> notes.memo 列のみ検索
                sql = "SELECT key FROM notes WHERE memo LIKE ?"
                params = (like_query_term,)

            else:  # is_db_global の場合
                # 'query' -> メモリ(df)が持つ列 + DBが持つ列 を検索

                # 1. まずDB側 (memo, full_text) を検索
                sql = (
                    "SELECT key FROM notes "
                    "WHERE memo LIKE ? OR full_text LIKE ?"
                )
                params = (like_query_term, like_query_term)

                cursor = db_conn.cursor()
                cursor.execute(sql, params)
                matching_keys = {row[0] for row in cursor.fetchall()}

                # 2. メモリ側 (df) の列も検索
                pandas_mask = (
                    df['title'].str.contains(
                        search_value, case=False, na=False) |
                    df['tags'].str.contains(
                        search_value, case=False, na=False) |
                    df['key'].str.contains(
                        search_value, case=False, na=False) |
                    df['time'].str.contains(
                        search_value, case=False, na=False) |
                    df['commonplace_key'].str.contains(
                        search_value, case=False) |
                    df['date'].str.contains(
                        search_value, case=False, na=False)
                )

                # 3. DBの結果(key)とメモリの結果(mask)を OR で結合
                term_condition = df['key'].isin(matching_keys) | pandas_mask

                logger.warning(
                    "[FTS] DB LIKE検索結果: "
                    f"{len(matching_keys)} 件の Key がヒット。",
                    extra={'sensitive': True}
                )
                return term_condition

            # (is_fts_column の場合、以下が実行される)
            cursor = db_conn.cursor()
            cursor.execute(sql, params)
            matching_keys = {row[0] for row in cursor.fetchall()}

            logger.warning(
                f"[FTS] DB LIKE検索結果: {len(matching_keys)} 件の Key がヒット。",
                extra={'sensitive': True}
            )

            # FTS結果とpandasの 'key' を比較し、boolマスクを返す
            term_condition = df['key'].isin(matching_keys)
            return term_condition

        except Exception as e:
            logger.warning(f"[FTS DEBUG] DB 'LIKE'検索エラー: {e}", exc_info=True)
            return pd.Series([False] * len(df), index=df.index)

    # --- FTS(LIKE)検索が実行されなかった場合 (Pandas検索) ---

    if term == search_value:  # グローバル検索かプレフィックス検索かを判定
        logger.warning(
            f"[FTS DEBUG] Pandas検索 (FTSスキップ) (Term: {term}, "
            f"is_db_col: {is_db_column}, is_db_global: {is_db_global}, "
            f"db_conn: {db_conn is not None})",
            extra={'sensitive': True}
        )

    if target_column == 'date':
        try:
            # YYYYMMDD-YYYYMMDD (範囲)
            range_match = re.match(r'^(\d{8})-(\d{8})$', search_value)
            # >=YYYYMMDD (以降)
            gte_match = re.match(r'^>=(\d{8})$', search_value)
            # <=YYYYMMDD (以前)
            lte_match = re.match(r'^<=(\d{8})$', search_value)

            # date列が文字列として比較可能であることを確認
            if 'date' not in df.columns:
                return pd.Series([False] * len(df), index=df.index)

            # 比較のために df['date'] を str 型に強制 (na=False は効かないため)
            date_series_str = df['date'].astype(str)

            if range_match:
                start_date = range_match.group(1)
                end_date = range_match.group(2)
                # .between() を使って範囲内の日付を検索
                term_condition = date_series_str.between(start_date, end_date)
            elif gte_match:
                start_date = gte_match.group(1)
                term_condition = date_series_str >= start_date
            elif lte_match:
                end_date = lte_match.group(1)
                term_condition = date_series_str <= end_date
            else:
                # 通常の部分一致検索 (例: date:202401)
                term_condition = date_series_str.str.contains(
                    search_value, case=False, na=False, regex=False
                )
        except Exception as e:
            logger.error(f"日付検索エラー: {e}")
            # エラー時は何もヒットしない (False) マスクを返す
            term_condition = pd.Series([False] * len(df), index=df.index)

    elif target_column:
        # --- プレフィックス検索 (日付以外): 指定された列のみ検索 ---
        # (target_column == 'memo' や 'full_text' は上で処理済み)
        if target_column == 'time':
            try:
                # 1. アンダースコア '_' を '.' (任意の一文字) に置換
                regex_pattern = search_value.replace('_', '.')

                # 2. 6桁未満の場合は、右側を '.' で6桁になるまで埋める
                if len(regex_pattern) < 6:
                    regex_pattern = regex_pattern.ljust(6, '.')

                # 3. 6桁を超える場合は、6桁に切り詰める
                elif len(regex_pattern) > 6:
                    regex_pattern = regex_pattern[:6]

                # 4. 完全一致の正規表現 ( ^...$ ) を作成
                final_regex = f"^{regex_pattern}$"
                logger.debug(f"[time: ] {final_regex} で検索")

                # 5. 正規表現検索を実行
                term_condition = df[target_column].str.contains(
                    final_regex,
                    case=False,  # (時刻なので case=False は不要だが念のため)
                    na=False,
                    regex=True   # ★ regex=True に設定
                )
            except Exception as e:
                logger.error(f"Time (Regex) 検索エラー: {e}")
                term_condition = pd.Series([False] * len(df), index=df.index)
        elif target_column in df.columns:
            # .str.contains() を使用して部分一致検索
            term_condition = df[target_column].str.contains(
                search_value, case=False, na=False, regex=False
            )
    else:
        # --- グローバル検索 (FTSが無効 or 本文検索OFF) ---
        term_condition = (
            df['title'].str.contains(search_value, case=False, na=False) |
            df['tags'].str.contains(search_value, case=False, na=False) |
            df['key'].str.contains(search_value, case=False, na=False) |
            df['time'].str.contains(search_value, case=False, na=False) |
            # (memo はFTS(LIKE)で検索されるため除外)
            df['commonplace_key'].str.contains(search_value, case=False) |
            df['date'].str.contains(search_value, case=False, na=False)
        )

    return term_condition


def parse_term(df, query, include_full_text=False, db_conn=None):
    """
    括弧、NOT(ハイフン)、または単純な検索語を処理する。

    Args:
        df (pd.DataFrame): 検索対象のDataFrame。
        query (str): 処理対象の検索クエリの一部 (例: 'title:A', '(A OR B)', '-C')。

    Returns:
        pd.Series: クエリに一致した行がTrueとなるboolマスク。
    """
    query = query.strip()
    is_not = False

    if query.startswith('-'):
        is_not = True
        query = query[1:].strip()

    if query.startswith('(') and query.endswith(')'):
        # 括弧の中身を評価するため、最上位のOR関数にフラグを渡して再帰呼び出し
        mask = parse_or_expression(df, query[1:-1], include_full_text, db_conn)
    else:
        # 最終的な検索実行関数にフラグを渡す
        mask = evaluate_simple_term(df, query, include_full_text, db_conn)

    return ~mask if is_not else mask


def parse_and_expression(df, query, include_full_text=False, db_conn=None):
    """
    AND 演算子で式を結合する。

    AND は OR よりも優先順位が高い。
    (例: 'A AND B AND C')

    Args:
        df (pd.DataFrame): 検索対象のDataFrame。
        query (str): ANDで結合された検索クエリ。

    Returns:
        pd.Series: クエリに一致した行がTrueとなるboolマスク。
    """
    and_parts = split_respecting_parens(query, ' AND ')

    # AND は「積」なので、Trueのマスクで初期化
    mask = pd.Series([True] * len(df), index=df.index)

    for part in and_parts:
        mask &= parse_term(df, part, include_full_text, db_conn)
    return mask


def parse_or_expression(df, query, include_full_text=False, db_conn=None):
    """
    OR 演算子で式を結合する (最上位の演算)。

    (例: '(A AND B) OR C')

    Args:
        df (pd.DataFrame): 検索対象のDataFrame。
        query (str): ORで結合された検索クエリ。

    Returns:
        pd.Series: クエリに一致した行がTrueとなるboolマスク。
    """
    or_parts = split_respecting_parens(query, ' OR ')

    # OR は「和」なので、Falseのマスクで初期化
    mask = pd.Series([False] * len(df), index=df.index)

    for part in or_parts:
        # 各パーツを AND 式として評価 (ANDが優先されるため)
        mask |= parse_and_expression(df, part, include_full_text, db_conn)
    return mask
