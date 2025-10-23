import pandas as pd


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


def evaluate_simple_term(df, term):
    """
    プレフィックス検索、またはグローバル検索を実行する。

    'title:Python' のようなプレフィックス検索、または 'Python' のような
    グローバル検索を処理し、該当する行のboolマスク (pd.Series) を返す。

    Args:
        df (pd.DataFrame): 検索対象のDataFrame。
        term (str): 単純な検索語 (例: 'Python', 'title:Python')。

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
            'ikey': 'commonplace_key'  # IndexKeyとその略称でも検索可
    }

    target_column = None
    final_search_term = term

    # プレフィックス (key:など) があるかチェック
    if ':' in term:
        parts = term.split(':', 1)
        prefix = parts[0].lower().strip()
        value = parts[1].strip()

        if prefix in search_fields_map and value:
            target_column = search_fields_map[prefix]
            final_search_term = value

    if not final_search_term:
        # 検索語が空なら、何もヒットしないマスクを返す
        return pd.Series([False] * len(df), index=df.index)

    term_condition = pd.Series([False] * len(df), index=df.index)

    if target_column:
        # --- プレフィックス検索: 指定された列のみ検索 ---
        if target_column in df.columns:
            # .str.contains() を使用して部分一致検索
            term_condition = df[target_column].str.contains(
                final_search_term, case=False, na=False, regex=False
            )
    else:
        # --- グローバル検索: 主要な列を検索 ---
        term_condition = (
            df['title'].str.contains(
                final_search_term, case=False, na=False, regex=False) |
            df['tags'].str.contains(
                final_search_term, case=False, na=False, regex=False) |
            df['key'].str.contains(
                final_search_term, case=False, na=False, regex=False) |
            df['memo'].str.contains(
                final_search_term, case=False, na=False, regex=False) |
            df['commonplace_key'].str.contains(
                final_search_term, case=False, na=False, regex=False) |
            df['date'].str.contains(
                final_search_term, case=False, na=False, regex=False)
        )
    return term_condition


def parse_term(df, query):
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
        # 括弧の中身を再帰的に評価 (最上位のORから)
        mask = parse_or_expression(df, query[1:-1])
    else:
        # 単純な検索語を評価
        mask = evaluate_simple_term(df, query)

    return ~mask if is_not else mask


def parse_and_expression(df, query):
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
        mask &= parse_term(df, part)
    return mask


def parse_or_expression(df, query):
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
        mask |= parse_and_expression(df, part)
    return mask
