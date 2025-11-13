import pandas as pd
import re  # 正規表現ライブラリをインポート

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


def evaluate_simple_term(df, term, include_full_text=False):
    """
    プレフィックス検索、またはグローバル検索を実行する。

    'title:Python' のようなプレフィックス検索、または 'Python' のような
    グローバル検索を処理し、該当する行のboolマスク (pd.Series) を返す。
    ['date'] プレフィックスの日時範囲検索 (>=, <=, YYYYMMDD-YYYYMMDD) にも対応。

    Args:
        df (pd.DataFrame): 検索対象のDataFrame。
        term (str): 単純な検索語 (例: 'Python', 'title:Python', 'date:>=20240101')。
        include_full_text (bool): グローバル検索時に本文も対象にするか。

    Returns:
        pd.Series: 検索条件に一致した行がTrueとなるboolマスク。
    """
    search_fields_map = {
            'title': 'title',
            'key': 'key',
            'date': 'date',
            'tag': 'tags',
            'tags': 'tags',  # 'tag'でも'tags'でも検索可
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

    # 日付検索ロジック
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
        if target_column in df.columns:
            # .str.contains() を使用して部分一致検索
            term_condition = df[target_column].str.contains(
                search_value, case=False, na=False, regex=False
            )
        # (もし target_column が 'full_text' であっても、
        #  df['full_text'] が検索されるだけで、include_full_text フラグは不要)

    else:
        # --- グローバル検索: 主要な列を検索 ---
        term_condition = (
            df['title'].str.contains(
                search_value, case=False, na=False, regex=False) |
            df['tags'].str.contains(
                search_value, case=False, na=False, regex=False) |
            df['key'].str.contains(
                search_value, case=False, na=False, regex=False) |
            df['memo'].str.contains(
                search_value, case=False, na=False, regex=False) |
            df['commonplace_key'].str.contains(
                search_value, case=False, na=False, regex=False) |
            df['date'].str.contains(
                search_value, case=False, na=False, regex=False)
        )

        # 「本文検索」が有効な場合、full_text カラムも検索対象に加える
        if include_full_text and 'full_text' in df.columns:
            term_condition |= df['full_text'].str.contains(
                search_value, case=False, na=False, regex=False
            )

    return term_condition


def parse_term(df, query, include_full_text=False):
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
        mask = parse_or_expression(df, query[1:-1], include_full_text)
    else:
        # 最終的な検索実行関数にフラグを渡す
        mask = evaluate_simple_term(df, query, include_full_text)

    return ~mask if is_not else mask


def parse_and_expression(df, query, include_full_text=False):
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
        mask &= parse_term(df, part, include_full_text)
    return mask


def parse_or_expression(df, query, include_full_text=False):
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
        mask |= parse_and_expression(df, part, include_full_text)
    return mask
