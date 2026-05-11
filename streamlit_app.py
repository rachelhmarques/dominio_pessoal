from __future__ import annotations

import hashlib
import json
from pathlib import Path
from tempfile import TemporaryDirectory

import pandas as pd
import streamlit as st

from dominio_parser import (
    START_LOT_OPTIONS,
    DomainParsingError,
    analyze_workbook,
    build_empty_mapping_rule,
    build_preview,
    merge_analysis_states,
    mapping_rules_from_state,
    serialize_preview_for_excel,
    source_options,
    write_csv_bytes,
)


st.set_page_config(
    page_title="Domínio Lançamento",
    page_icon="📊",
    layout="wide",
)


APP_DIR = Path(__file__).parent
SHARED_MAPPING_PATH = APP_DIR / "data" / "shared_mapping_rules.json"
GLOBAL_MAPPING_KEY = "__global__"
ANALYSIS_SIGNATURE_VERSION = "2026-05-07-branch67-v2"
SPECIALIZED_BRANCH_KEYS = {"67"}


def main() -> None:
    inject_styles()
    init_session_state()
    render_header()
    handle_upload()

    if not st.session_state.analysis_state:
        render_empty_state()
        return

    sync_rules_from_editor()
    preview = current_preview()
    render_summary(preview)
    render_exports(preview)
    render_entries(preview)
    render_memory(preview)
    render_mapping(preview)


def inject_styles() -> None:
    st.markdown(
        """
        <style>
        .stApp {
            background:
                radial-gradient(circle at top left, rgba(15,118,110,.12), transparent 34%),
                radial-gradient(circle at bottom right, rgba(161,98,7,.08), transparent 30%),
                linear-gradient(135deg, #f7f0e4 0%, #eadcc7 100%);
        }
        .hero {
            padding: 1.4rem 1.6rem;
            border: 1px solid rgba(217, 200, 174, .9);
            border-radius: 24px;
            background: rgba(255, 251, 244, .86);
            box-shadow: 0 24px 60px rgba(70, 56, 29, .12);
            margin-bottom: 1rem;
        }
        .eyebrow {
            color: #0f766e;
            text-transform: uppercase;
            letter-spacing: .14em;
            font-size: .75rem;
            font-weight: 700;
            margin-bottom: .5rem;
        }
        .hero h1 {
            margin: 0;
            font-size: 2.8rem;
            line-height: 1;
            color: #21302e;
        }
        .hero p {
            color: #5e6c66;
            font-size: 1rem;
            line-height: 1.65;
            max-width: 70ch;
            margin-top: .9rem;
            margin-bottom: 0;
        }
        .section-card {
            padding: 1rem 1.1rem;
            border: 1px solid rgba(217, 200, 174, .9);
            border-radius: 22px;
            background: rgba(255, 251, 244, .82);
            box-shadow: 0 16px 40px rgba(70, 56, 29, .08);
            margin-bottom: 1rem;
        }
        .metric-card {
            padding: 1rem;
            border-radius: 18px;
            background: rgba(255,255,255,.72);
            border: 1px solid rgba(217, 200, 174, .9);
        }
        .metric-card .label {
            color: #5e6c66;
            text-transform: uppercase;
            letter-spacing: .08em;
            font-size: .78rem;
            font-weight: 700;
        }
        .metric-card .value {
            color: #21302e;
            font-size: 1.9rem;
            font-weight: 700;
            line-height: 1.1;
            margin-top: .35rem;
        }
        .helper {
            color: #5e6c66;
            font-size: .92rem;
            line-height: 1.55;
        }
        .stDownloadButton > button, .stButton > button {
            border-radius: 999px !important;
            font-weight: 700 !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def init_session_state() -> None:
    st.session_state.setdefault("analysis_state", None)
    st.session_state.setdefault("uploaded_signature", None)
    st.session_state.setdefault("mapping_rules", [])
    st.session_state.setdefault("mapping_editor_snapshot", None)


def load_shared_mapping_store() -> dict[str, list[dict[str, object]]]:
    if not SHARED_MAPPING_PATH.exists():
        return {}
    try:
        payload = json.loads(SHARED_MAPPING_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}
    if not isinstance(payload, dict):
        return {}

    normalized: dict[str, list[dict[str, object]]] = {}
    for branch_code, rules in payload.items():
        if not isinstance(branch_code, str) or not isinstance(rules, list):
            continue
        valid_rules = [dict(rule) for rule in rules if isinstance(rule, dict)]
        if valid_rules:
            normalized[branch_code] = valid_rules
    return normalized


def save_shared_mapping_store(store: dict[str, list[dict[str, object]]]) -> None:
    SHARED_MAPPING_PATH.parent.mkdir(parents=True, exist_ok=True)
    SHARED_MAPPING_PATH.write_text(
        json.dumps(store, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def normalize_rule_list(rules: list[dict[str, object]]) -> list[dict[str, object]]:
    return [rule.to_dict() for rule in mapping_rules_from_state({"mapping_rules": rules})]


def merge_mapping_rules(
    default_rules: list[dict[str, object]],
    saved_rules: list[dict[str, object]],
) -> list[dict[str, object]]:
    normalized_defaults = normalize_rule_list(default_rules)
    normalized_saved = normalize_rule_list(saved_rules)
    default_by_id = {
        str(rule["rule_id"]): dict(rule)
        for rule in normalized_defaults
    }

    merged: list[dict[str, object]] = []
    seen_rule_ids: set[str] = set()

    for saved_rule in normalized_saved:
        rule_id = str(saved_rule["rule_id"])
        if rule_id in default_by_id:
            merged_rule = dict(default_by_id[rule_id])
            merged_rule.update(saved_rule)
            merged.append(merged_rule)
        else:
            merged.append(dict(saved_rule))
        seen_rule_ids.add(rule_id)

    for default_rule in normalized_defaults:
        rule_id = str(default_rule["rule_id"])
        if rule_id not in seen_rule_ids:
            merged.append(dict(default_rule))

    return merged


def saved_mapping_rules_for_branch(branch_code: str) -> list[dict[str, object]] | None:
    store = load_shared_mapping_store()
    rules = store.get(branch_code)
    if not rules:
        return None
    return normalize_rule_list(rules)


def saved_global_mapping_rules() -> list[dict[str, object]] | None:
    store = load_shared_mapping_store()
    rules = store.get(GLOBAL_MAPPING_KEY)
    if rules:
        return normalize_rule_list(rules)
    return None


def saved_mapping_rules_to_apply(analyzed_state: dict[str, object]) -> list[dict[str, object]] | None:
    branch_code = str(analyzed_state.get("branch_code") or "")
    if branch_code:
        branch_rules = saved_mapping_rules_for_branch(branch_code)
        if branch_rules:
            return branch_rules
    if branch_code in SPECIALIZED_BRANCH_KEYS:
        return None
    return saved_global_mapping_rules()


def resolve_mapping_rules(analyzed_state: dict[str, object]) -> list[dict[str, object]]:
    default_rules = [dict(rule) for rule in list(analyzed_state["mapping_rules"])]
    saved_rules = saved_mapping_rules_to_apply(analyzed_state)
    if not saved_rules:
        return default_rules
    return merge_mapping_rules(default_rules, saved_rules)


def render_header() -> None:
    st.markdown(
        """
        <div class="hero">
          <div class="eyebrow">Importação Setor Pessoal</div>
          <h1>Domínio Lançamento em Streamlit</h1>
          <p>
            Envie o <strong>Resumo Mensal</strong> e, se houver, o arquivo complementar do
            <strong>13º</strong>. Depois confira os lançamentos gerados na tela, revise a
            memória de cálculo, ajuste o mapeamento e exporte o resultado em
            <strong>CSV</strong> ou <strong>Excel</strong>.
          </p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def handle_upload() -> None:
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    uploaded_file = st.file_uploader(
        "Arquivo principal do Resumo Mensal",
        type=["xls", "xlsx"],
        help="No Streamlit Cloud, prefira arquivos .xls ou .xlsx lidos diretamente em Python.",
    )
    thirteenth_file = st.file_uploader(
        "Arquivo complementar do 13º (opcional)",
        type=["xls", "xlsx"],
        help="Use aqui o modelo Resumo Mensal13 quando quiser consolidar o 13º no mesmo resultado.",
    )

    if uploaded_file is not None:
        file_bytes = uploaded_file.getvalue()
        secondary_bytes = thirteenth_file.getvalue() if thirteenth_file is not None else b""
        signature = hashlib.sha256(
            file_bytes
            + b"::"
            + secondary_bytes
            + b"::"
            + ANALYSIS_SIGNATURE_VERSION.encode("utf-8")
        ).hexdigest()
        if signature != st.session_state.uploaded_signature:
            try:
                with TemporaryDirectory() as temp_dir:
                    temp_path = Path(temp_dir) / uploaded_file.name
                    temp_path.write_bytes(file_bytes)
                    analyzed_state = analyze_workbook(temp_path)

                    if thirteenth_file is not None:
                        thirteenth_path = Path(temp_dir) / thirteenth_file.name
                        thirteenth_path.write_bytes(secondary_bytes)
                        analyzed_state = merge_analysis_states(
                            [analyzed_state, analyze_workbook(thirteenth_path)]
                        )
            except DomainParsingError as exc:
                st.error(str(exc))
                st.session_state.analysis_state = None
                st.session_state.mapping_rules = []
                st.session_state.mapping_editor_snapshot = None
            except Exception as exc:
                st.error(f"Não foi possível processar a planilha: {exc}")
                st.session_state.analysis_state = None
                st.session_state.mapping_rules = []
                st.session_state.mapping_editor_snapshot = None
            else:
                st.session_state.analysis_state = analyzed_state
                st.session_state.mapping_rules = resolve_mapping_rules(analyzed_state)
                st.session_state.mapping_editor_snapshot = None
                st.session_state.uploaded_signature = signature

    st.markdown(
        '<div class="helper">A lógica de cálculo fica separada da interface, então o mapeamento pode ser ajustado sem reler as planilhas a cada edição.</div>',
        unsafe_allow_html=True,
    )
    st.markdown("</div>", unsafe_allow_html=True)


def current_state() -> dict[str, object] | None:
    base_state = st.session_state.analysis_state
    if not base_state:
        return None
    state = dict(base_state)
    state["mapping_rules"] = list(st.session_state.mapping_rules)
    return state


def current_preview() -> dict[str, object]:
    state = current_state()
    if state is None:
        raise RuntimeError("Tentativa de montar prévia sem estado de análise.")
    return build_preview(state)


def save_current_mapping_as_shared_default() -> None:
    store = load_shared_mapping_store()
    store[GLOBAL_MAPPING_KEY] = normalize_rule_list(list(st.session_state.mapping_rules))
    save_shared_mapping_store(store)


def restore_shared_mapping() -> bool:
    base_state = st.session_state.analysis_state
    if not base_state:
        return False

    saved_rules = saved_mapping_rules_to_apply(base_state)
    if not saved_rules:
        return False

    st.session_state.mapping_rules = merge_mapping_rules(
        list(base_state["mapping_rules"]),
        saved_rules,
    )
    st.session_state.mapping_editor_snapshot = None
    return True


def reset_mapping_to_analysis_defaults() -> None:
    base_state = st.session_state.analysis_state
    if not base_state:
        return
    st.session_state.mapping_rules = [dict(rule) for rule in list(base_state["mapping_rules"])]
    st.session_state.mapping_editor_snapshot = None


def render_empty_state() -> None:
    st.info("Envie o Resumo Mensal e, se necessário, o arquivo complementar do 13º para gerar a prévia dos lançamentos.")


def render_summary(preview: dict[str, object]) -> None:
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("Resumo da análise")
    caption_label = "Arquivos" if " + " in str(preview["workbook_name"]) else "Arquivo"
    st.caption(f"{caption_label} {preview['workbook_name']} · Filial {preview['branch_code']}")

    cols = st.columns(4)
    metrics = [
        ("Filial", str(preview["branch_code"])),
        ("Competências", str(preview["summary"]["competencies"])),
        ("Lançamentos", str(preview["summary"]["entries"])),
        ("Total", str(preview["summary"]["total_value"])),
    ]
    for column, (label, value) in zip(cols, metrics):
        column.markdown(
            f"""
            <div class="metric-card">
              <div class="label">{label}</div>
              <div class="value">{value}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    st.markdown("</div>", unsafe_allow_html=True)


def render_exports(preview: dict[str, object]) -> None:
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    left, right = st.columns([1.2, 1])
    with left:
        st.subheader("Exportações")
        st.markdown(
            '<div class="helper">A exportação usa a configuração de mapeamento atual da tela.</div>',
            unsafe_allow_html=True,
        )
    with right:
        col_csv, col_xlsx = st.columns(2)
        csv_bytes = write_csv_bytes(list(preview["entry_csv_rows"]))
        excel_bytes = serialize_preview_for_excel(preview)
        suffix = " (com 13º)" if preview.get("has_thirteenth_summary") else ""
        col_csv.download_button(
            "Baixar CSV",
            data=csv_bytes,
            file_name=f"DOMINIO LANCAMENTO - {preview['branch_code']}{suffix}.csv",
            mime="text/csv",
            use_container_width=True,
        )
        col_xlsx.download_button(
            "Baixar Excel",
            data=excel_bytes,
            file_name=f"DOMINIO LANCAMENTO - {preview['branch_code']}{suffix}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    st.markdown("</div>", unsafe_allow_html=True)


def render_entries(preview: dict[str, object]) -> None:
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("Lançamentos gerados")
    st.markdown(
        '<div class="helper">Cada linha mostra a origem do valor e a memória de cálculo correspondente.</div>',
        unsafe_allow_html=True,
    )
    if preview.get("needs_thirteenth_summary"):
        st.warning(
            "O arquivo atual indica movimentos de 13º, mas o fechamento anual do 13º não foi enviado. "
            "Para gerar os lançamentos anuais e os reflexos extras de FGTS/PIS em dezembro, envie também o arquivo complementar `Resumo Mensal13`."
        )
    entries_df = pd.DataFrame(list(preview["entries"]))
    if not entries_df.empty:
        entries_df = entries_df.rename(
            columns={
                "index": "#",
                "competency_reference": "Competência",
                "mapping_label": "Mapeamento",
                "source_label": "Fonte",
                "posting_date": "Data",
                "debit_account": "Débito",
                "credit_account": "Crédito",
                "value": "Valor",
                "history": "Histórico",
                "start_lot": "Inicia lote",
                "memory": "Memória de cálculo",
            }
        )
        st.dataframe(entries_df, use_container_width=True, hide_index=True)
    st.markdown("</div>", unsafe_allow_html=True)


def render_memory(preview: dict[str, object]) -> None:
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("Memória por competência")
    st.markdown(
        '<div class="helper">As fontes calculadas abaixo alimentam o mapeamento e os lançamentos.</div>',
        unsafe_allow_html=True,
    )

    for competency in list(preview["competencies"]):
        with st.expander(f"{competency['reference']} · lançamento em {competency['posting_date']}", expanded=False):
            sources_df = pd.DataFrame(list(competency["sources"]))
            if not sources_df.empty:
                sources_df = sources_df.rename(
                    columns={
                        "label": "Fonte",
                        "description": "Descrição",
                        "value": "Valor",
                        "memory": "Memória de cálculo",
                    }
                )
                st.dataframe(sources_df[["Fonte", "Descrição", "Valor", "Memória de cálculo"]], use_container_width=True, hide_index=True)
            else:
                st.info("Nenhuma fonte positiva encontrada para esta competência.")
    st.markdown("</div>", unsafe_allow_html=True)


def render_mapping(preview: dict[str, object]) -> None:
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    title_col, button_col = st.columns([1.6, 1])
    with title_col:
        st.subheader("Mapeamento editável")
        st.markdown(
            '<div class="helper">Edite contas, histórico, ordem, estratégia de lote e adicione novas regras. O preview acima atualiza automaticamente. Quando salvo, esse padrão passa a valer para todas as empresas.</div>',
            unsafe_allow_html=True,
        )
    with button_col:
        if st.button("Adicionar mapeamento", use_container_width=True):
            add_mapping_rule()
            st.rerun()

    st.caption("Placeholders suportados no histórico: {reference}, {branch_code}, {month}, {year}, {posting_date}, {source_label}")

    current_rules = list(st.session_state.mapping_rules)
    rule_ids = [str(rule["rule_id"]) for rule in current_rules]
    rules_df = pd.DataFrame(
        [
            {
                "Ativo": bool(rule["active"]),
                "Ordem": int(rule["order"]),
                "Rótulo": str(rule["label"]),
                "Fonte": str(rule["source_key"]),
                "Conta débito": str(rule["debit_account"]),
                "Conta crédito": str(rule["credit_account"]),
                "Histórico": str(rule["history_template"]),
                "Inicia lote": str(rule["start_lot_strategy"]),
            }
            for rule in current_rules
        ]
    )

    source_keys = [str(option["key"]) for option in source_options()]
    editor = st.data_editor(
        rules_df,
        key="mapping_editor",
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        column_config={
            "Ativo": st.column_config.CheckboxColumn("Ativo"),
            "Ordem": st.column_config.NumberColumn("Ordem", step=1, format="%d"),
            "Rótulo": st.column_config.TextColumn("Rótulo"),
            "Fonte": st.column_config.SelectboxColumn("Fonte", options=source_keys),
            "Conta débito": st.column_config.TextColumn("Conta débito"),
            "Conta crédito": st.column_config.TextColumn("Conta crédito"),
            "Histórico": st.column_config.TextColumn("Histórico"),
            "Inicia lote": st.column_config.SelectboxColumn(
                "Inicia lote",
                options=list(START_LOT_OPTIONS.keys()),
            ),
        },
    )

    st.session_state.mapping_rules = dataframe_to_mapping_rules(editor, rule_ids)
    st.session_state.mapping_editor_snapshot = editor.to_dict(orient="records")

    action_col, restore_col, reset_col = st.columns(3)
    if action_col.button("Salvar padrão global", use_container_width=True):
        save_current_mapping_as_shared_default()
        st.success("Mapeamento global salvo para próximas sessões e para todas as empresas.")
    if restore_col.button("Recarregar padrão global", use_container_width=True):
        if restore_shared_mapping():
            st.rerun()
        st.info("Nenhum padrão global salvo foi encontrado.")
    if reset_col.button("Voltar ao padrão da análise", use_container_width=True):
        reset_mapping_to_analysis_defaults()
        st.rerun()

    if saved_global_mapping_rules():
        st.caption(f"Há um padrão global compartilhado salvo em `{SHARED_MAPPING_PATH.name}`.")
    else:
        st.caption("Ainda não existe padrão global compartilhado salvo.")
    st.markdown("</div>", unsafe_allow_html=True)


def dataframe_to_mapping_rules(dataframe: pd.DataFrame, rule_ids: list[str]) -> list[dict[str, object]]:
    mapping_rules: list[dict[str, object]] = []
    for row_id, row in zip(rule_ids, dataframe.to_dict(orient="records")):
        mapping_rules.append(
            {
                "rule_id": row_id,
                "label": str(row.get("Rótulo", "")).strip() or "Mapeamento sem nome",
                "source_key": str(row.get("Fonte", "regular_salary")).strip() or "regular_salary",
                "debit_account": str(row.get("Conta débito", "")).strip(),
                "credit_account": str(row.get("Conta crédito", "")).strip(),
                "history_template": str(row.get("Histórico", "")).strip() or "{source_label} {reference}",
                "start_lot_strategy": str(row.get("Inicia lote", "never")).strip() or "never",
                "active": bool(row.get("Ativo", True)),
                "order": safe_int(row.get("Ordem", 0)),
            }
        )
    return mapping_rules


def normalize_editor_snapshot(
    snapshot: object,
    fallback_rows: list[dict[str, object]],
) -> list[dict[str, object]] | None:
    if isinstance(snapshot, pd.DataFrame):
        return snapshot.to_dict(orient="records")

    if isinstance(snapshot, list):
        normalized_rows: list[dict[str, object]] = []
        for row in snapshot:
            if isinstance(row, dict):
                normalized_rows.append(dict(row))
        return normalized_rows

    if not isinstance(snapshot, dict):
        return None

    if all(isinstance(key, str) for key in snapshot.keys()) and "edited_rows" not in snapshot:
        try:
            dataframe = pd.DataFrame(snapshot)
        except ValueError:
            return None
        return dataframe.to_dict(orient="records")

    rows = [dict(row) for row in fallback_rows]
    edited_rows = snapshot.get("edited_rows", {})
    if isinstance(edited_rows, dict):
        for row_index, changes in edited_rows.items():
            try:
                index = int(row_index)
            except Exception:
                continue
            if not isinstance(changes, dict) or not (0 <= index < len(rows)):
                continue
            rows[index].update(changes)

    added_rows = snapshot.get("added_rows", [])
    if isinstance(added_rows, list):
        for row in added_rows:
            if isinstance(row, dict):
                rows.append(dict(row))

    deleted_rows = snapshot.get("deleted_rows", [])
    if isinstance(deleted_rows, list):
        for row_index in sorted(
            (safe_int(row) for row in deleted_rows),
            reverse=True,
        ):
            if 0 <= row_index < len(rows):
                rows.pop(row_index)

    return rows


def safe_int(value: object) -> int:
    try:
        return int(value)
    except Exception:
        return 0


def add_mapping_rule() -> None:
    current_rules = mapping_rules_from_state(current_state() or {"mapping_rules": []})
    next_order = max((rule.order for rule in current_rules), default=0) + 10
    current_rules.append(build_empty_mapping_rule(next_order))
    st.session_state.mapping_rules = [rule.to_dict() for rule in current_rules]


def sync_rules_from_editor() -> None:
    snapshot = st.session_state.get("mapping_editor")
    cached = st.session_state.mapping_editor_snapshot
    if snapshot and snapshot != cached and st.session_state.mapping_rules:
        fallback_rows = cached if isinstance(cached, list) else []
        normalized_rows = normalize_editor_snapshot(snapshot, fallback_rows)
        if normalized_rows is None:
            return
        rule_ids = [str(rule["rule_id"]) for rule in st.session_state.mapping_rules]
        dataframe = pd.DataFrame(normalized_rows)
        st.session_state.mapping_rules = dataframe_to_mapping_rules(dataframe, rule_ids)
        st.session_state.mapping_editor_snapshot = normalized_rows


if __name__ == "__main__":
    main()
