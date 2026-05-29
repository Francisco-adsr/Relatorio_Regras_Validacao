import re
import pyodbc
import pandas as pd
from dataclasses import dataclass
from typing import List, Optional, Dict

# ============================================================
# CONFIGURAÇÕES
# ============================================================

CONNECTION_STRING = (
    "DRIVER={SQL Server};"
    "SERVER=NB-CB2RB93;"
    "DATABASE=GDMS_Central_Vale_Base_Metals;"
    "UID=admin;"
    "PWD=fus10nAtValeNF;"
    "TrustServerCertificate=yes;"
)

RULES_TABLE = "VALIDATION_RULES"
RULE_NAME_COL = "RULE_NAME"
RULE_SQL_COL = "RULE_SQL"
BASE_TABLE_COL = "RULE_SQL_FROM_TABLE"
ENABLE_RULE_COL = "ENABLE_RULE"

DRILLHOLE_TABLE = "DRILL_HOLE"
DRILLHOLE_KEY = "HOLE_NUMBER"
PROJECT_KEY = "PROJECT_NUMBER"

INPUT_HOLES_FILE = "furos_para_validar.csv"
INPUT_HOLE_COLUMN = "HOLE_NUMBER"

INPUT_RULES_FILE = "regras_para_validar.csv"
INPUT_RULE_COLUMN = "RULE_NAME"

OUTPUT_FILE = "relatorio_validacao_furos_03.xlsx"
DETAIL_OUTPUT_FILE = "relatorio_validacao_por_regra_03.xlsx"
BOOLEAN_OUTPUT_FILE = "relatorio_validacao_matriz_regras_03.xlsx"

# ============================================================
# MODELO
# ============================================================

@dataclass
class ValidationRule:
    rule_name: str
    base_table: str
    rule_sql: str
    enable_rule: str

# ============================================================
# ENTRADA
# ============================================================

def fetch_input_holes(csv_file: str) -> pd.DataFrame:
    df = pd.read_csv(csv_file)

    if INPUT_HOLE_COLUMN not in df.columns:
        raise ValueError(
            f"O arquivo CSV deve conter a coluna '{INPUT_HOLE_COLUMN}'."
        )

    df = df[[INPUT_HOLE_COLUMN]].copy()
    df[INPUT_HOLE_COLUMN] = df[INPUT_HOLE_COLUMN].astype(str).str.strip()
    df = df[df[INPUT_HOLE_COLUMN] != ""]
    df = df.drop_duplicates().sort_values(INPUT_HOLE_COLUMN).reset_index(drop=True)

    df = df.rename(columns={INPUT_HOLE_COLUMN: "HOLE_NUMBER"})
    return df

def fetch_input_rules(csv_file: str) -> List[str]:
    df = pd.read_csv(csv_file)

    if INPUT_RULE_COLUMN not in df.columns:
        raise ValueError(
            f"O arquivo CSV de regras deve conter a coluna '{INPUT_RULE_COLUMN}'."
        )

    df = df[[INPUT_RULE_COLUMN]].copy()
    df[INPUT_RULE_COLUMN] = df[INPUT_RULE_COLUMN].astype(str).str.strip()
    df = df[df[INPUT_RULE_COLUMN] != ""]
    df = df.drop_duplicates().sort_values(INPUT_RULE_COLUMN).reset_index(drop=True)

    return df[INPUT_RULE_COLUMN].tolist()

# ============================================================
# PROMPT
# ============================================================

def ask_rule_selection_mode() -> str:
    print("\nSelecione o modo de execução das regras:")
    print("1 - Usar o CSV de entrada com a lista de regras que deseja validar")
    print("2 - Rodar todas as regras da tabela VALIDATION_RULES")
    print("3 - Rodar apenas as regras com ENABLE_RULE = 'Y'")

    while True:
        option = input("Digite 1, 2 ou 3: ").strip()
        if option in {"1", "2", "3"}:
            return option
        print("Opção inválida. Digite apenas 1, 2 ou 3.")

# ============================================================
# UTILITÁRIOS SQL
# ============================================================

def normalize_sql(sql: str) -> str:
    """Remove comentários e normaliza espaços."""
    sql = re.sub(r"/\*.*?\*/", "", sql, flags=re.S)
    sql = re.sub(r"--.*?$", "", sql, flags=re.M)
    sql = sql.strip().rstrip(";")
    return sql

def validate_full_rule_sql(sql: str) -> Optional[str]:
    """
    Valida a SQL completa da regra.
    Bloqueia comandos perigosos e estruturas que não sejam SELECT.
    """
    normalized = normalize_sql(sql)
    upper_sql = normalized.upper()

    forbidden = ["DROP ", "DELETE ", "INSERT ", "UPDATE ", "EXEC ", "MERGE ", "TRUNCATE "]
    if any(token in upper_sql for token in forbidden):
        return "Regra contém comando proibido. Apenas SELECT é permitido."

    if not upper_sql.startswith("SELECT") and not upper_sql.startswith("WITH"):
        return "A RULE_SQL deve começar com SELECT ou WITH."

    return None

def extract_where_clause(sql: str) -> Optional[str]:
    sql = normalize_sql(sql)

    match = re.search(r"\bWHERE\b\s+(.+)$", sql, flags=re.I | re.S)
    if not match:
        return None

    where_clause = match.group(1).strip()
    where_clause = re.sub(r"\bORDER\s+BY\b.+$", "", where_clause, flags=re.I | re.S).strip()

    return where_clause

def extract_columns_from_where(where_clause: str, valid_aliases: set) -> List[dict]:
    if not where_clause:
        return []

    clean_where = remove_subqueries_from_where(where_clause)

    pattern = r"((?:\[[^\]]+\]|\b[A-Za-z_][A-Za-z0-9_]*\b)\s*\.\s*(?:\[[^\]]+\]|\b[A-Za-z_][A-Za-z0-9_]*\b))"
    matches = re.findall(pattern, clean_where, flags=re.I)

    results = []
    seen = set()

    for match in matches:
        expression = re.sub(r"\s+", "", match)
        parts = expression.split(".")
        if len(parts) != 2:
            continue

        alias = parts[0].strip("[] ")
        col = parts[1].strip("[] ")

        if valid_aliases and alias.upper() not in valid_aliases:
            continue

        if col.upper() == "HOLE_NUMBER":
            continue

        key = col.upper()
        if key in seen:
            continue

        seen.add(key)
        results.append({
            "expression": expression,
            "alias": col
        })

    return results

def extract_outer_query_aliases(rule_sql: str, base_table: str) -> set:
    main_alias = extract_main_table_alias(rule_sql, base_table)
    table_name = base_table.split(".")[-1].strip("[] ")
    return {main_alias.upper(), table_name.upper()}

def remove_subqueries_from_where(where_clause: str) -> str:
    if not where_clause:
        return where_clause

    text = where_clause

    pattern = re.compile(r"\(\s*SELECT\b", flags=re.I)

    while True:
        match = pattern.search(text)
        if not match:
            break

        start = match.start()
        depth = 0
        end = None

        for i in range(start, len(text)):
            if text[i] == "(":
                depth += 1
            elif text[i] == ")":
                depth -= 1
                if depth == 0:
                    end = i
                    break

        if end is None:
            break

        text = text[:start] + " " + text[end + 1:]

    return text

def sanitize_sheet_name(name: str, used_names: set) -> str:
    """
    Excel limita nome de aba a 31 caracteres e proíbe certos caracteres.
    """
    cleaned = re.sub(r'[:\\/*?\[\]]', "_", name)
    cleaned = cleaned[:31].strip()

    if not cleaned:
        cleaned = "REGRA"

    original = cleaned
    counter = 1
    while cleaned in used_names:
        suffix = f"_{counter}"
        cleaned = (original[:31 - len(suffix)] + suffix).strip()
        counter += 1

    used_names.add(cleaned)
    return cleaned

def normalize_dataframe_columns(df: pd.DataFrame) -> pd.DataFrame:
    result = df.copy()
    result.columns = [str(col).strip() for col in result.columns]
    return result

def quote_identifier(name: str) -> str:
    """
    Protege identificador SQL Server com colchetes.
    Ex: tabela.schema -> [tabela].[schema]
    """
    parts = [p.strip("[] ") for p in name.split(".")]
    return ".".join(f"[{p}]" for p in parts if p)

def build_sql_in_list(values: List[str]) -> str:
    escaped = []
    for v in values:
        v_escaped = v.replace("'", "''")
        escaped.append(f"'{v_escaped}'")
    return ", ".join(escaped)

def create_temp_holes_table(conn, holes: List[str]):
    cursor = conn.cursor()

    cursor.execute("""
    IF OBJECT_ID('tempdb..#TEMP_HOLES') IS NOT NULL
        DROP TABLE #TEMP_HOLES;

    CREATE TABLE #TEMP_HOLES (
        HOLE_NUMBER VARCHAR(100) COLLATE DATABASE_DEFAULT
    );
    """)

    insert_sql = "INSERT INTO #TEMP_HOLES (HOLE_NUMBER) VALUES (?)"

    cursor.fast_executemany = True
    cursor.executemany(insert_sql, [(h,) for h in holes])

    conn.commit()

def filter_existing_holes(input_holes_df: pd.DataFrame, db_holes_df: pd.DataFrame):
    input_df = input_holes_df.copy()
    db_df = db_holes_df.copy()

    input_df["HOLE_NUMBER"] = input_df["HOLE_NUMBER"].astype(str).str.strip()
    db_df["HOLE_NUMBER"] = db_df["HOLE_NUMBER"].astype(str).str.strip()

    if "PROJECT" not in db_df.columns:
        db_df["PROJECT"] = ""

    db_df["PROJECT"] = db_df["PROJECT"].fillna("").astype(str).str.strip()

    valid_holes_df = input_df.merge(
        db_df[["HOLE_NUMBER", "PROJECT"]].drop_duplicates(),
        on="HOLE_NUMBER",
        how="inner"
    ).drop_duplicates().sort_values("HOLE_NUMBER").reset_index(drop=True)

    invalid_holes_df = input_df.merge(
        db_df[["HOLE_NUMBER"]].drop_duplicates(),
        on="HOLE_NUMBER",
        how="left",
        indicator=True
    )

    invalid_holes_df = invalid_holes_df[invalid_holes_df["_merge"] == "left_only"][["HOLE_NUMBER"]]
    invalid_holes_df["STATUS"] = "NOT_FOUND_IN_DRILL_HOLE"
    invalid_holes_df = invalid_holes_df.sort_values("HOLE_NUMBER").reset_index(drop=True)

    return valid_holes_df, invalid_holes_df

def get_table_alias(base_table: str) -> str:
    return base_table.split(".")[-1].strip("[] ")

def extract_main_table_alias(rule_sql: str, base_table: str) -> str:
    """
    Retorna o alias real da tabela base dentro da RULE_SQL.
    Se a tabela base não tiver alias explícito, retorna o nome da tabela.

    Exemplos:
      FROM dbo.TABELA t       -> retorna "t"
      FROM [dbo].[TABELA] t   -> retorna "t"
      FROM TABELA             -> retorna "TABELA"
      FROM TABELA WHERE ...   -> retorna "TABELA"
    """
    sql = normalize_sql(rule_sql)

    table_parts = [p.strip("[] ") for p in base_table.split(".") if p.strip()]
    table_name = table_parts[-1]
    schema_name = table_parts[-2] if len(table_parts) > 1 else None

    table_name_pattern = rf"\[?{re.escape(table_name)}\]?"
    schema_table_pattern = table_name_pattern

    if schema_name:
        schema_name_pattern = rf"\[?{re.escape(schema_name)}\]?"
        schema_table_pattern = rf"{schema_name_pattern}\s*\.\s*{table_name_pattern}"

    patterns = [
        # FROM schema.tabela AS alias
        rf"\bFROM\s+{schema_table_pattern}\s+AS\s+([A-Za-z_][A-Za-z0-9_]*)\b",

        # FROM schema.tabela alias
        rf"\bFROM\s+{schema_table_pattern}\s+([A-Za-z_][A-Za-z0-9_]*)\b",

        # FROM tabela AS alias
        rf"\bFROM\s+{table_name_pattern}\s+AS\s+([A-Za-z_][A-Za-z0-9_]*)\b",

        # FROM tabela alias
        rf"\bFROM\s+{table_name_pattern}\s+([A-Za-z_][A-Za-z0-9_]*)\b",
    ]

    reserved_words = {
        "WHERE", "JOIN", "INNER", "LEFT", "RIGHT", "FULL", "ON",
        "GROUP", "ORDER", "HAVING", "UNION", "EXCEPT", "INTERSECT"
    }

    for pattern in patterns:
        match = re.search(pattern, sql, flags=re.I)
        if match:
            alias = match.group(1).strip()
            if alias.upper() not in reserved_words:
                return alias

    return table_name

def extract_scalar_subqueries(where_clause: str) -> List[str]:
    """
    Extrai subqueries escalares do WHERE no formato (SELECT ...).
    Retorna a expressão completa com parênteses.
    """
    if not where_clause:
        return []

    text = where_clause
    pattern = re.compile(r"\(\s*SELECT\b", flags=re.I)

    results = []
    start_pos = 0

    while True:
        match = pattern.search(text, start_pos)
        if not match:
            break

        start = match.start()
        depth = 0
        end = None

        for i in range(start, len(text)):
            if text[i] == "(":
                depth += 1
            elif text[i] == ")":
                depth -= 1
                if depth == 0:
                    end = i
                    break

        if end is None:
            break

        subquery = text[start:end + 1].strip()
        results.append(subquery)
        start_pos = end + 1

    return results

def extract_scalar_subqueries_from_where(where_clause: str) -> List[str]:
    """
    Extrai apenas subqueries escalares do WHERE.
    Ignora EXISTS(...) e NOT EXISTS(...).
    """
    if not where_clause:
        return []

    text = where_clause
    results = []

    pattern = re.compile(
        r"(=|<>|!=|<|>|<=|>=)\s*(\(\s*SELECT\b)",
        flags=re.I
    )

    start_pos = 0
    while True:
        match = pattern.search(text, start_pos)
        if not match:
            break

        paren_start = match.start(2)
        depth = 0
        end = None

        for i in range(paren_start, len(text)):
            if text[i] == "(":
                depth += 1
            elif text[i] == ")":
                depth -= 1
                if depth == 0:
                    end = i
                    break

        if end is None:
            break

        subquery = text[paren_start:end + 1].strip()
        results.append(subquery)
        start_pos = end + 1

    return results

def derive_alias_from_scalar_subquery(subquery: str) -> str:
    """
    Tenta derivar um alias amigável para a subquery com base
    na primeira expressão do SELECT interno.

    Exemplos:
      (SELECT MIN(GEO.description_date) FROM ...) -> description_date
      (SELECT GEO.description_date FROM ...)      -> description_date
      (SELECT MAX([TAB].[XRF_DATE]) FROM ...)     -> XRF_DATE
    """
    inner = subquery.strip()

    if inner.startswith("(") and inner.endswith(")"):
        inner = inner[1:-1].strip()

    # captura o trecho entre SELECT e FROM
    match = re.search(
        r"^\s*SELECT\s+(?:DISTINCT\s+)?(.+?)\s+FROM\b",
        inner,
        flags=re.I | re.S
    )
    if not match:
        return "SUBQUERY_VALUE"

    select_expr = match.group(1).strip()

    # remove alias explícito do tipo "... AS nome"
    explicit_alias = re.search(
        r"\bAS\s+\[?([A-Za-z_][A-Za-z0-9_]*)\]?\s*$",
        select_expr,
        flags=re.I
    )
    if explicit_alias:
        return explicit_alias.group(1)

    # tenta pegar o último identificador coluna de algo como:
    # MIN(GEO.description_date), GEO.description_date, [GEO].[description_date]
    col_matches = re.findall(
        r"(?:\[[^\]]+\]|\b[A-Za-z_][A-Za-z0-9_]*\b)\s*\.\s*(?:\[[^\]]+\]|\b[A-Za-z_][A-Za-z0-9_]*\b)",
        select_expr,
        flags=re.I
    )
    if col_matches:
        last_ref = col_matches[-1]
        return last_ref.split(".")[-1].strip("[] ")

    # fallback: limpa função/agregação e usa nome genérico
    simple_id = re.search(r"\b([A-Za-z_][A-Za-z0-9_]*)\b\s*$", select_expr)
    if simple_id:
        return simple_id.group(1)

    return "SUBQUERY_VALUE"

def extract_exists_predicates(where_clause: str) -> List[dict]:
    """
    Extrai predicados EXISTS(...) e NOT EXISTS(...) do WHERE.
    Retorna uma lista de dicts com:
      - predicate: expressão completa
      - alias: nome sugerido para a coluna booleana
      - kind: exists_flag
    """
    if not where_clause:
        return []

    text = where_clause
    results = []
    counter_exists = 1
    counter_not_exists = 1

    pattern = re.compile(r"\b(NOT\s+EXISTS|EXISTS)\s*\(\s*SELECT\b", flags=re.I)

    start_pos = 0
    while True:
        match = pattern.search(text, start_pos)
        if not match:
            break

        keyword = match.group(1).upper()
        start = match.start()

        paren_start = text.find("(", match.start())
        if paren_start == -1:
            break

        depth = 0
        end = None
        for i in range(paren_start, len(text)):
            if text[i] == "(":
                depth += 1
            elif text[i] == ")":
                depth -= 1
                if depth == 0:
                    end = i
                    break

        if end is None:
            break

        predicate = text[start:end + 1].strip()

        if keyword == "EXISTS":
            alias = f"EXISTS_{counter_exists}"
            counter_exists += 1
        else:
            alias = f"NOT_EXISTS_{counter_not_exists}"
            counter_not_exists += 1

        results.append({
            "expression": f"CASE WHEN {predicate} THEN 1 ELSE 0 END",
            "alias": alias,
            "kind": "exists_flag"
        })

        start_pos = end + 1

    return results

def extract_required_expressions_from_where(where_clause: str, valid_aliases: set) -> List[dict]:
    """
    Retorna expressões úteis para o SELECT detalhado:
    - colunas externas do WHERE
    - subqueries escalares
    Ignora EXISTS / NOT EXISTS.
    """
    results = []
    seen = set()

    # Colunas da query externa
    outer_columns = extract_columns_from_where(where_clause, valid_aliases)
    for col in outer_columns:
        alias_upper = col["alias"].strip("[] ").upper()
        if alias_upper not in seen:
            results.append(col)
            seen.add(alias_upper)

    # Subqueries escalares
    scalar_subqueries = extract_scalar_subqueries_from_where(where_clause)
    for subquery in scalar_subqueries:
        alias = derive_alias_from_scalar_subquery(subquery)
        alias_upper = alias.strip("[] ").upper()

        if alias_upper not in seen:
            results.append({
                "expression": subquery,
                "alias": alias
            })
            seen.add(alias_upper)

    return results

def inject_required_columns_into_select(rule_sql: str, base_table: str) -> str:
    sql = normalize_sql(rule_sql)
    main_alias = extract_main_table_alias(sql, base_table)

    select_match = re.search(
        r"^\s*SELECT\s+(DISTINCT\s+)?(TOP\s+\(?\d+\)?\s+)?",
        sql,
        flags=re.I
    )
    if not select_match:
        raise ValueError("Não foi possível localizar o SELECT principal da RULE_SQL.")

    select_list_start = select_match.end()

    from_match = re.search(r"\bFROM\b", sql[select_list_start:], flags=re.I)
    if not from_match:
        raise ValueError("Não foi possível localizar o FROM da RULE_SQL.")

    from_pos = select_list_start + from_match.start()

    select_prefix = sql[:select_list_start]
    select_list = sql[select_list_start:from_pos]
    rest_sql = sql[from_pos:]

    where_clause = extract_where_clause(sql)
    valid_aliases = extract_outer_query_aliases(sql, base_table)

    required_columns = [
        {"expression": f"{main_alias}.HOLE_NUMBER", "alias": "HOLE_NUMBER"}
    ]

    where_expressions = extract_required_expressions_from_where(where_clause, valid_aliases)

    seen_required = {"HOLE_NUMBER"}
    for item in where_expressions:
        alias_upper = item["alias"].strip("[] ").upper()
        if alias_upper not in seen_required:
            required_columns.append(item)
            seen_required.add(alias_upper)             
                   
    existing_aliases = set()

    # captura aliases explícitos: "... AS NOME"
    explicit_aliases = re.findall(
        r"\bAS\s+\[?([A-Za-z_][A-Za-z0-9_]*)\]?",
        select_list,
        flags=re.I
    )
    existing_aliases.update(a.upper() for a in explicit_aliases)

    # captura referências simples já presentes no select: alias.coluna
    simple_cols = re.findall(
        r"(?:\[[^\]]+\]|\b[A-Za-z_][A-Za-z0-9_]*\b)\s*\.\s*(?:\[[^\]]+\]|\b[A-Za-z_][A-Za-z0-9_]*\b)",
        select_list,
        flags=re.I
    )
    for col in simple_cols:
        col_name = col.split(".")[-1].strip("[] ").upper()
        existing_aliases.add(col_name)

    columns_to_inject = []
    injected_aliases = set()

    for col in required_columns:
        alias = col["alias"].strip("[] ").upper()
        if alias in existing_aliases or alias in injected_aliases:
            continue

        columns_to_inject.append(f"{col['expression']} AS {col['alias']}")
        injected_aliases.add(alias)

    if not columns_to_inject:
        return sql

    new_select_list = ", ".join(columns_to_inject) + ", " + select_list.strip()

    return f"{select_prefix}{new_select_list} {rest_sql}".strip()

def build_full_rule_query(rule_sql: str, base_table: str, rule_name: str) -> str:
    """
    Usa a RULE_SQL completa, injeta HOLE_NUMBER no SELECT,
    encapsula em subquery e filtra apenas os furos da #TEMP_HOLES.
    """
    sql_error = validate_full_rule_sql(rule_sql)
    if sql_error:
        raise ValueError(sql_error)

    sql_with_required_columns = inject_required_columns_into_select(rule_sql, base_table)

    final_sql = f"""
    SELECT DISTINCT
        r.*
    FROM (
        {sql_with_required_columns}
    ) r
    INNER JOIN #TEMP_HOLES t
        ON t.HOLE_NUMBER COLLATE DATABASE_DEFAULT =
        r.HOLE_NUMBER COLLATE DATABASE_DEFAULT
    """.strip()

    return final_sql

def attach_project_column(df: pd.DataFrame, holes_df: pd.DataFrame) -> pd.DataFrame:
    """
    Adiciona a coluna PROJECT a qualquer DataFrame que contenha HOLE_NUMBER.
    """
    if df.empty or "HOLE_NUMBER" not in df.columns:
        return df

    result = df.copy()
    result.columns = [str(col).strip() for col in result.columns]

    if "PROJECT" in result.columns:
        result = result.drop(columns=["PROJECT"])
    
    result["HOLE_NUMBER"] = result["HOLE_NUMBER"].astype(str).str.strip()

    lookup = holes_df[["HOLE_NUMBER", "PROJECT"]].drop_duplicates().copy()
    lookup["HOLE_NUMBER"] = lookup["HOLE_NUMBER"].astype(str).str.strip()
    lookup["PROJECT"] = lookup["PROJECT"].astype(str).str.strip()

    result = result.merge(lookup, on="HOLE_NUMBER", how="left")

    cols = result.columns.tolist()
    if "PROJECT" in cols and "HOLE_NUMBER" in cols:
        cols.remove("PROJECT")
        hole_idx = cols.index("HOLE_NUMBER")
        cols.insert(hole_idx + 1, "PROJECT")
        result = result[cols]

    return result

# ============================================================
# BANCO
# ============================================================

def get_connection():
    return pyodbc.connect(CONNECTION_STRING)

def fetch_rules(conn) -> List[ValidationRule]:
    sql = f"""
    SELECT
        {RULE_NAME_COL} AS rule_name,
        {BASE_TABLE_COL} AS base_table,
        {RULE_SQL_COL} AS rule_sql,
        COALESCE({ENABLE_RULE_COL}, 'N') AS enable_rule
    FROM {RULES_TABLE}
    WHERE {RULE_NAME_COL} IS NOT NULL
      AND {BASE_TABLE_COL} IS NOT NULL
      AND {RULE_SQL_COL} IS NOT NULL
      AND {BASE_TABLE_COL} NOT LIKE '%SSTN%'
    ORDER BY {RULE_NAME_COL}
    """

    df = pd.read_sql(sql, conn)

    return [
        ValidationRule(
            rule_name=str(row["rule_name"]).strip(),
            base_table=str(row["base_table"]).strip(),
            rule_sql=str(row["rule_sql"]).strip(),
            enable_rule=str(row["enable_rule"] or "N").strip().upper()
        )
        for _, row in df.iterrows()
    ]

def fetch_all_holes(conn) -> pd.DataFrame:
    sql = f"""
    SELECT DISTINCT 
        {DRILLHOLE_KEY} AS HOLE_NUMBER,
        {PROJECT_KEY} AS PROJECT
    FROM {DRILLHOLE_TABLE}
    WHERE {DRILLHOLE_KEY} IS NOT NULL
    ORDER BY {DRILLHOLE_KEY}
    """
    df = pd.read_sql(sql, conn)
    df["HOLE_NUMBER"] = df["HOLE_NUMBER"].astype(str).str.strip()

    if "PROJECT" in df.columns:
        df["PROJECT"] = df["PROJECT"].fillna("").astype(str).str.strip()
    else:
        df["PROJECT"] = ""

    df = df[df["HOLE_NUMBER"] != ""].drop_duplicates().reset_index(drop=True)
    return df

# ============================================================
# EXECUÇÃO DAS REGRAS
# ============================================================

def filter_rules_by_mode(
    rules: List[ValidationRule],
    mode: str,
    selected_rule_names: Optional[List[str]] = None
) -> List[ValidationRule]:

    if mode == "1":
        if not selected_rule_names:
            raise ValueError(
                f"Nenhuma regra foi informada no arquivo {INPUT_RULES_FILE}."
            )

        selected_set = {
            rule_name.strip().upper()
            for rule_name in selected_rule_names
            if str(rule_name).strip()
        }

        filtered_rules = [
            rule for rule in rules
            if rule.rule_name.upper() in selected_set
        ]

        missing_rules = sorted(
            selected_set - {rule.rule_name.upper() for rule in filtered_rules}
        )

        if missing_rules:
            print("Aviso: as seguintes regras do CSV não foram encontradas na VALIDATION_RULES:")
            for rule_name in missing_rules:
                print(f" - {rule_name}")

        if not filtered_rules:
            raise ValueError(
                "Nenhuma regra do CSV foi encontrada na tabela VALIDATION_RULES."
            )

        return filtered_rules

    elif mode == "2":
        return rules

    elif mode == "3":
        filtered_rules = [rule for rule in rules if rule.enable_rule == "Y"]

        if not filtered_rules:
            raise ValueError("Nenhuma regra ativa com ENABLE_RULE = 'Y' foi encontrada.")

        return filtered_rules

    else:
        raise ValueError("Modo de seleção de regras inválido.")

def execute_rules(conn, rules: List[ValidationRule], selected_holes_df: pd.DataFrame):
    """
    Executa as regras de validação usando a RULE_SQL completa.

    Retorna:

    - failures_by_rule:
        dict em que a chave é o nome da regra
        e o valor é um DataFrame com:
            HOLE_NUMBER | PROJECT | VALIDATION_RULE | demais colunas originais da regra

    - errors_df:
        DataFrame com regras que falharam na validação ou execução

    - all_failures_df:
        DataFrame consolidado mínimo com:
            HOLE_NUMBER | VALIDATION_RULE
    """
    selected_holes = selected_holes_df["HOLE_NUMBER"].dropna().astype(str).tolist()
    create_temp_holes_table(conn, selected_holes)
    failures_by_rule: Dict[str, pd.DataFrame] = {}
    errors = []
    all_failures = []

    if not selected_holes:
        raise ValueError("Nenhum furo informado para validação.")

    hole_list_sql = build_sql_in_list(selected_holes)


    for rule in rules:
        query = None
        try:
            # 1) Normaliza a SQL completa da regra
            normalized_sql = normalize_sql(rule.rule_sql)

            # 2) Valida a SQL completa
            sql_error = validate_full_rule_sql(normalized_sql)
            if sql_error:
                errors.append({
                    "RULE_NAME": rule.rule_name,
                    "BASE_TABLE": rule.base_table,
                    "ERROR": sql_error
                })
                continue

            # 3) Monta a query final usando a RULE_SQL inteira
            query = build_full_rule_query(
                rule_sql=normalized_sql,
                base_table=rule.base_table,
                rule_name=rule.rule_name
            )

            print("\n====================================================")
            print(f"RULE_NAME: {rule.rule_name}")
            print("FINAL SQL:")
            print(query)
            print("====================================================\n")

            # 4) Executa a query
            df = pd.read_sql(query, conn)
            df = normalize_dataframe_columns(df)

            # Normaliza espaços dos nomes das colunas
            df.columns = [str(col).strip() for col in df.columns]

            # Cria mapa case-insensitive das colunas retornadas
            column_map_upper = {str(col).strip().upper(): col for col in df.columns}

            # Garante que HOLE_NUMBER exista, independentemente de maiúsculas/minúsculas
            if "HOLE_NUMBER" not in column_map_upper:
                errors.append({
                    "RULE_NAME": rule.rule_name,
                    "BASE_TABLE": rule.base_table,
                    "ERROR": f"A query final não retornou HOLE_NUMBER. Colunas retornadas: {list(df.columns)}",
                    "FINAL_SQL": query
                })
                continue

            # Renomeia a coluna encontrada para o padrão esperado
            real_hole_col = column_map_upper["HOLE_NUMBER"]
            if real_hole_col != "HOLE_NUMBER":
                df = df.rename(columns={real_hole_col: "HOLE_NUMBER"})

            # 6) Normaliza HOLE_NUMBER
            df["HOLE_NUMBER"] = df["HOLE_NUMBER"].astype(str).str.strip()
            df = df[df["HOLE_NUMBER"] != ""]

            # 7) Se não retornou nada, guarda dataframe vazio
            if df.empty:
                failures_by_rule[rule.rule_name] = df
                continue

            # 8) Adiciona PROJECT
            df = attach_project_column(df, selected_holes_df)

            # 9) Opcional: adiciona nome da regra para rastreabilidade
            if "VALIDATION_RULE" not in df.columns:
                df["VALIDATION_RULE"] = rule.rule_name

            # 10) Reordena colunas para deixar HOLE_NUMBER e PROJECT na frente
            priority_cols = [col for col in ["HOLE_NUMBER", "PROJECT", "VALIDATION_RULE"] if col in df.columns]
            other_cols = [col for col in df.columns if col not in priority_cols]
            df = df[priority_cols + other_cols]

            # 11) Remove duplicidades
            df = df.drop_duplicates().reset_index(drop=True)

            # 12) Guarda o detalhado completo por regra
            failures_by_rule[rule.rule_name] = df

            # 13) Monta consolidado mínimo para os demais relatórios
            df_fail = df[["HOLE_NUMBER"]].drop_duplicates().copy()
            df_fail["VALIDATION_RULE"] = rule.rule_name

            if not df_fail.empty:
                all_failures.append(df_fail)

        except Exception as e:
            errors.append({
                "RULE_NAME": rule.rule_name,
                "BASE_TABLE": rule.base_table,
                "ERROR": str(e),
                "FINAL_SQL": query
            })

    # 14) Consolida todas as falhas
    if all_failures:
        all_failures_df = pd.concat(all_failures, ignore_index=True).drop_duplicates()
        all_failures_df = all_failures_df.sort_values(
            ["HOLE_NUMBER", "VALIDATION_RULE"]
        ).reset_index(drop=True)
    else:
        all_failures_df = pd.DataFrame(columns=["HOLE_NUMBER", "VALIDATION_RULE"])

    # 15) Consolida erros
    errors_df = pd.DataFrame(errors)

    return failures_by_rule, errors_df, all_failures_df

def build_authorized_df(all_holes_df: pd.DataFrame, all_failures_df: pd.DataFrame) -> pd.DataFrame:
    all_holes_df = all_holes_df.copy()
    all_holes_df["HOLE_NUMBER"] = all_holes_df["HOLE_NUMBER"].astype(str).str.strip()

    if "PROJECT" not in all_holes_df.columns:
        all_holes_df["PROJECT"] = ""

    all_holes_df["PROJECT"] = all_holes_df["PROJECT"].fillna("").astype(str).str.strip()

    if all_failures_df.empty:
        result = all_holes_df[["HOLE_NUMBER", "PROJECT"]].copy()
        result["STATUS"] = "OK"
        return result

    all_failures_df = all_failures_df.copy()
    all_failures_df["HOLE_NUMBER"] = all_failures_df["HOLE_NUMBER"].astype(str).str.strip()

    failed_holes = all_failures_df[["HOLE_NUMBER"]].drop_duplicates()

    result = all_holes_df.merge(
        failed_holes,
        on="HOLE_NUMBER",
        how="left",
        indicator=True
    )

    result = result[result["_merge"] == "left_only"][["HOLE_NUMBER", "PROJECT"]].copy()
    result["STATUS"] = "OK"
    result = result.sort_values("HOLE_NUMBER").reset_index(drop=True)

    return result

def build_summary_df(failures_by_rule: dict, errors_df: pd.DataFrame) -> pd.DataFrame:
    rows = []

    for rule_name, df in failures_by_rule.items():
        rows.append({
            "RULE_NAME": rule_name,
            "FAILED_HOLES": len(df),
            "STATUS": "OK"
        })

    if not errors_df.empty:
        for _, row in errors_df.iterrows():
            rows.append({
                "RULE_NAME": row["RULE_NAME"],
                "FAILED_HOLES": None,
                "STATUS": f"ERROR: {row['ERROR']}"
            })

    summary_df = pd.DataFrame(rows)

    if not summary_df.empty:
        summary_df = summary_df.sort_values(["STATUS", "FAILED_HOLES", "RULE_NAME"], ascending=[True, False, True])

    return summary_df

def build_failed_holes_report(all_failures_df: pd.DataFrame, valid_holes_df: pd.DataFrame) -> pd.DataFrame:
    """
    Gera um relatório em que cada linha é um HOLE_NUMBER
    e as regras falhadas aparecem em colunas:
    HOLE_NUMBER | VALIDATION_RULE_1 | VALIDATION_RULE_2 | VALIDATION_RULE_3 | ...
    """
    if all_failures_df.empty:
        return pd.DataFrame(columns=["HOLE_NUMBER", "PROJECT"])

    df = all_failures_df.copy()
    df = df.sort_values(["HOLE_NUMBER", "VALIDATION_RULE"]).reset_index(drop=True)

    # Numera as regras dentro de cada furo
    df["RULE_ORDER"] = df.groupby("HOLE_NUMBER").cumcount() + 1

    # Faz o pivot: cada regra vira uma coluna
    pivot_df = df.pivot(
        index="HOLE_NUMBER",
        columns="RULE_ORDER",
        values="VALIDATION_RULE"
    )

    # Renomeia as colunas para VALIDATION_RULE_1, VALIDATION_RULE_2, VALIDATION_RULE_3...
    pivot_df.columns = [f"VALIDATION_RULE_{col}" for col in pivot_df.columns]
    pivot_df = pivot_df.reset_index()

    pivot_df = attach_project_column(pivot_df, valid_holes_df)

    return pivot_df

def build_rule_boolean_report(
    valid_holes_df: pd.DataFrame,
    rules: List[ValidationRule],
    failures_by_rule: dict
) -> pd.DataFrame:
    result = valid_holes_df[["HOLE_NUMBER", "PROJECT"]].drop_duplicates().copy()
    result["HOLE_NUMBER"] = result["HOLE_NUMBER"].astype(str).str.strip()
    result["PROJECT"] = result["PROJECT"].astype(str).str.strip()
    result = result.sort_values("HOLE_NUMBER").reset_index(drop=True)

    for rule in rules:
        rule_name = rule.rule_name

        failed_df = failures_by_rule.get(rule_name, pd.DataFrame(columns=["HOLE_NUMBER"]))

        failed_holes = set(
            failed_df["HOLE_NUMBER"].astype(str).str.strip().dropna().tolist()
        )

        result[rule_name] = result["HOLE_NUMBER"].isin(failed_holes)

    return result

# ============================================================
# EXPORTAÇÃO EXCEL
# ============================================================

def export_to_excel(
    output_file: str,
    summary_df: pd.DataFrame,
    authorized_df: pd.DataFrame,
    failed_holes_report_df: pd.DataFrame,
    errors_df: pd.DataFrame
):
    used_sheet_names = set()

    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
        workbook = writer.book

        header_fmt = workbook.add_format({
            "bold": True,
            "text_wrap": False
        })

        # RESUMO
        summary_sheet = sanitize_sheet_name("RESUMO", used_sheet_names)
        summary_df.to_excel(writer, sheet_name=summary_sheet, index=False)
        ws = writer.sheets[summary_sheet]
        for col_num, col_name in enumerate(summary_df.columns):
            ws.write(0, col_num, col_name, header_fmt)
        ws.set_column(0, max(len(summary_df.columns) - 1, 0), 28)

        # AUTORIZADOS
        auth_sheet = sanitize_sheet_name("AUTORIZADOS", used_sheet_names)
        authorized_df.to_excel(writer, sheet_name=auth_sheet, index=False)
        ws = writer.sheets[auth_sheet]
        for col_num, col_name in enumerate(authorized_df.columns):
            ws.write(0, col_num, col_name, header_fmt)
        ws.set_column(0, max(len(authorized_df.columns) - 1, 0), 22)

        # FALHAS POR FURO
        failed_sheet = sanitize_sheet_name("FALHAS_POR_FURO", used_sheet_names)
        failed_holes_report_df.to_excel(writer, sheet_name=failed_sheet, index=False)
        ws = writer.sheets[failed_sheet]
        for col_num, col_name in enumerate(failed_holes_report_df.columns):
            ws.write(0, col_num, col_name, header_fmt)
        ws.set_column(0, max(len(failed_holes_report_df.columns) - 1, 0), 28)

        # ERROS
        if not errors_df.empty:
            err_sheet = sanitize_sheet_name("ERROS", used_sheet_names)
            errors_df.to_excel(writer, sheet_name=err_sheet, index=False)
            ws = writer.sheets[err_sheet]
            for col_num, col_name in enumerate(errors_df.columns):
                ws.write(0, col_num, col_name, header_fmt)
            ws.set_column(0, max(len(errors_df.columns) - 1, 0), 40)

def export_rule_details_to_excel(
    output_file: str,
    failures_by_rule: dict,
    valid_holes_df: pd.DataFrame
):
    used_sheet_names = set()

    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
        workbook = writer.book

        header_fmt = workbook.add_format({
            "bold": True,
            "text_wrap": False
        })

        rules_without_failures = []

        for rule_name, df in failures_by_rule.items():
            if df.empty:
                rules_without_failures.append(rule_name)
                continue

            sheet_name = sanitize_sheet_name(rule_name, used_sheet_names)

            rule_df = df.copy()

            if "PROJECT" not in rule_df.columns:
                rule_df = attach_project_column(rule_df, valid_holes_df)

            if "HOLE_NUMBER" in rule_df.columns:
                rule_df["HOLE_NUMBER"] = rule_df["HOLE_NUMBER"].astype(str).str.strip()

            priority_cols = [col for col in ["HOLE_NUMBER", "PROJECT", "VALIDATION_RULE"] if col in rule_df.columns]
            other_cols = [col for col in rule_df.columns if col not in priority_cols]
            rule_df = rule_df[priority_cols + other_cols]

            if "HOLE_NUMBER" in rule_df.columns:
                rule_df = rule_df.sort_values("HOLE_NUMBER").reset_index(drop=True)
            else:
                rule_df = rule_df.reset_index(drop=True)

            rule_df.to_excel(writer, sheet_name=sheet_name, index=False)

            ws = writer.sheets[sheet_name]
            for col_num, col_name in enumerate(rule_df.columns):
                ws.write(0, col_num, col_name, header_fmt)

            ws.set_column(0, len(rule_df.columns) - 1, 25)

        no_fail_df = pd.DataFrame({
            "RULE_NAME": sorted(rules_without_failures)
        })

        no_fail_sheet = sanitize_sheet_name("REGRAS_SEM_FALHAS", used_sheet_names)
        no_fail_df.to_excel(writer, sheet_name=no_fail_sheet, index=False)

        ws = writer.sheets[no_fail_sheet]
        for col_num, col_name in enumerate(no_fail_df.columns):
            ws.write(0, col_num, col_name, header_fmt)

        ws.set_column(0, 0, 40)

def export_boolean_report_to_excel(
    output_file: str,
    boolean_report_df: pd.DataFrame
):
    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
        workbook = writer.book

        header_fmt = workbook.add_format({
            "bold": True,
            "text_wrap": False
        })

        sheet_name = "MATRIZ_REGRAS"
        boolean_report_df.to_excel(writer, sheet_name=sheet_name, index=False)

        ws = writer.sheets[sheet_name]

        for col_num, col_name in enumerate(boolean_report_df.columns):
            ws.write(0, col_num, col_name, header_fmt)

        ws.set_column(0, 0, 25)   # HOLE_NUMBER
        if len(boolean_report_df.columns) > 1:
            ws.set_column(1, len(boolean_report_df.columns) - 1, 18)

# ============================================================
# MAIN
# ============================================================

def main():
    conn = None
    try:
        conn = get_connection()

        print("Lendo regras...")
        all_rules = fetch_rules(conn)
        print(f"Regras encontradas na VALIDATION_RULES: {len(all_rules)}")

        rule_selection_mode = ask_rule_selection_mode()
        selected_rule_names = None

        if rule_selection_mode == "1":
            print(f"Lendo regras informadas no CSV: {INPUT_RULES_FILE}")
            selected_rule_names = fetch_input_rules(INPUT_RULES_FILE)
            print(f"Regras informadas no CSV: {len(selected_rule_names)}")

        rules = filter_rules_by_mode(
            rules=all_rules,
            mode=rule_selection_mode,
            selected_rule_names=selected_rule_names
        )
        print(f"Regras selecionadas para execução: {len(rules)}")

        print("Lendo furos...")
        input_holes_df = fetch_input_holes(INPUT_HOLES_FILE)
        print(f"Furos informados no arquivo: {len(input_holes_df)}")

        print("Buscando furos existentes no banco...")
        db_holes_df = fetch_all_holes(conn)
        print(f"Furos existentes na DRILL_HOLE: {len(db_holes_df)}")

        print("Validando furos de entrada contra a DRILL_HOLE...")
        valid_holes_df, invalid_holes_df = filter_existing_holes(input_holes_df, db_holes_df)
        print(f"Furos válidos para processamento: {len(valid_holes_df)}")
        print(f"Furos não encontrados na DRILL_HOLE: {len(invalid_holes_df)}")

        if valid_holes_df.empty:
            raise ValueError("Nenhum HOLE_NUMBER do arquivo de entrada existe na tabela DRILL_HOLE.")

        print("Executando regras...")
        failures_by_rule, errors_df, all_failures_df = execute_rules(conn, rules, valid_holes_df)

        print("Montando autorizados...")
        authorized_df = build_authorized_df(valid_holes_df, all_failures_df)

        print("Montando resumo...")
        summary_df = build_summary_df(failures_by_rule, errors_df)
        failed_holes_report_df = build_failed_holes_report(all_failures_df, valid_holes_df)

        print("Montando matriz booleana por regra...")
        boolean_report_df = build_rule_boolean_report(
            valid_holes_df, 
            rules, 
            failures_by_rule
        )

        print("Exportando Excel...")
        export_to_excel(
            OUTPUT_FILE,
            summary_df,
            authorized_df,
            failed_holes_report_df,
            errors_df
        )

        print("Exportando Excel detalhado por regra...")
        export_rule_details_to_excel(
            DETAIL_OUTPUT_FILE,
            failures_by_rule,
            valid_holes_df
        )

        print("Exportando Excel matriz booleana por regra...")
        export_boolean_report_to_excel(
            BOOLEAN_OUTPUT_FILE,
            boolean_report_df
        )

        print("Concluído com sucesso.")
        print(f"Arquivo gerado: {OUTPUT_FILE}")
        print(f"Arquivo gerado: {DETAIL_OUTPUT_FILE}")
        print(f"Arquivo gerado: {BOOLEAN_OUTPUT_FILE}")
        print(f"Regras processadas: {len(rules)}")
        print(f"Regras com erro: {len(errors_df)}")
        print(f"Furos autorizados: {len(authorized_df)}")
        print(f"Ocorrências totais de falha: {len(all_failures_df)}")

    except Exception as e:
        print(f"Erro na execução: {e}")
        raise
    finally:
        if conn is not None:
            conn.close()

if __name__ == "__main__":
    main()