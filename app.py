import io
import os
from pathlib import Path

import pandas as pd
import streamlit as st
from dotenv import load_dotenv

load_dotenv(Path(__file__).resolve().parent / ".env", encoding="utf-8")

st.set_page_config(page_title="Préparer un fichier", layout="wide")

NEEDED = ("COL_ID", "COL_SITE", "COL_DOC", "COL_ACTOR", "COL_LABEL", "MAX_KEEP", "REASON_ACTOR", "REASON_MAX")


def cfg(name):
    value = None
    try:
        if name in st.secrets and str(st.secrets[name]).strip():
            value = str(st.secrets[name]).strip()
    except Exception:
        value = None
    if not value:
        raw = os.getenv(name)
        if raw and raw.strip():
            value = raw.strip()
    return value or None


def load_settings():
    found = {name: cfg(name) for name in NEEDED}
    if any(value is None for value in found.values()):
        return None
    columns = [found[name] for name in ("COL_ID", "COL_SITE", "COL_DOC", "COL_ACTOR", "COL_LABEL")]
    if len(set(columns)) != len(columns):
        return None
    try:
        found["MAX_KEEP"] = int(found["MAX_KEEP"])
    except ValueError:
        return None
    if found["MAX_KEEP"] < 1:
        return None
    return found


def classify(df, settings):
    group_a = settings["COL_ID"]
    group_b = settings["COL_SITE"]
    order_a = settings["COL_DOC"]
    order_b = settings["COL_ACTOR"]
    ordered = df.sort_values([group_a, group_b, order_a, order_b], kind="mergesort")
    first = (
        ordered.dropna(subset=[order_b])
        .groupby([group_a, group_b])[order_b]
        .transform("first")
    )
    same = ordered[order_b].eq(first).reindex(df.index).fillna(False)
    ranked = ordered[ordered[order_b].eq(first)].copy()
    ranked["_n"] = ranked.groupby([group_a, group_b]).cumcount() + 1
    keep = ranked.index[ranked["_n"] <= settings["MAX_KEEP"]]

    marked = df.copy()
    marked["_keep"] = False
    marked.loc[keep, "_keep"] = True
    marked["_why"] = ""
    blocked = ~marked["_keep"]
    marked.loc[blocked & ~same, "_why"] = settings["REASON_ACTOR"]
    marked.loc[blocked & same, "_why"] = settings["REASON_MAX"]
    return marked


def read_matching_sheet(file_bytes, columns):
    try:
        book = pd.ExcelFile(io.BytesIO(file_bytes))
    except Exception:
        return None
    for sheet in book.sheet_names:
        try:
            table = pd.read_excel(book, sheet_name=sheet)
        except Exception:
            continue
        if set(columns).issubset(table.columns):
            return table
    return "format"


@st.cache_data(show_spinner=False)
def prepare(file_bytes, packed):
    settings = dict(zip(NEEDED, packed))
    settings["MAX_KEEP"] = int(settings["MAX_KEEP"])
    columns = [settings[name] for name in ("COL_ID", "COL_SITE", "COL_DOC", "COL_ACTOR", "COL_LABEL")]
    table = read_matching_sheet(file_bytes, columns)
    if table is None or isinstance(table, str):
        return table
    marked = classify(table, settings)
    clean = marked.loc[marked["_keep"], table.columns]
    view = marked.drop(columns=["_keep"])
    view.insert(0, "Motif", view.pop("_why"))
    view.insert(0, "Conservé", marked["_keep"].map({True: "Oui", False: "Non"}))
    return view, to_excel_bytes(clean)


def to_excel_bytes(table):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        table.to_excel(writer, index=False, sheet_name="Sheet1")
    return buffer.getvalue()


def ask_password(expected):
    if not expected or st.session_state.get("access_granted"):
        return
    st.title("Préparer un fichier")
    with st.form("access"):
        entered = st.text_input("Mot de passe", type="password")
        submitted = st.form_submit_button("Entrer")
    if submitted:
        if entered == expected:
            st.session_state["access_granted"] = True
            st.rerun()
        st.error("Accès refusé.")
    st.stop()


def main():
    settings = load_settings()
    if settings is None:
        st.error("La configuration de l'application est incomplète.")
        st.stop()

    ask_password(cfg("APP_PASSWORD"))

    st.title("Préparer un fichier")

    uploaded = st.file_uploader("Fichier Excel", type=["xlsx"])
    if uploaded is None:
        st.stop()

    file_marker = f"{uploaded.name}:{uploaded.size}"
    if st.session_state.get("file_marker") != file_marker:
        st.session_state["file_marker"] = file_marker
        st.session_state.pop("client_choice", None)

    packed = tuple(str(settings[name]) for name in NEEDED)
    with st.spinner("Lecture du fichier..."):
        prepared = prepare(uploaded.getvalue(), packed)

    if prepared is None:
        st.error("Le fichier n'a pas pu être lu.")
        st.stop()
    if prepared == "format":
        st.error("Ce fichier ne peut pas être traité.")
        st.stop()

    view, clean_file = prepared
    if view.empty:
        st.warning("Le fichier ne contient aucune ligne.")
        st.stop()

    kept = int((view["Conservé"] == "Oui").sum())
    removed = int((view["Conservé"] == "Non").sum())
    total_col, kept_col, removed_col = st.columns(3)
    total_col.metric("Lignes", len(view))
    kept_col.metric("Conservées", kept)
    removed_col.metric("Retirées", removed)

    stem = Path(uploaded.name).stem
    st.download_button(
        "Télécharger le fichier conservé",
        data=clean_file,
        file_name=f"{stem}_retenu.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

    st.divider()
    st.subheader("Voir le détail")

    label = settings["COL_LABEL"]
    client_names = view[label].where(view[label].notna(), "(sans nom)").astype(str).str.strip()
    client_names = client_names.mask(client_names.eq(""), "(sans nom)")
    with_removed = client_names[view["Conservé"] == "Non"]
    choices = sorted(with_removed.unique(), key=str.lower)
    if not choices:
        st.info("Aucune ligne retirée.")
        st.stop()

    only_removed = st.checkbox("Afficher seulement les lignes retirées")
    chosen = st.selectbox("Choisir", choices, key="client_choice")

    shown = view.loc[client_names == chosen]
    if only_removed:
        shown = shown.loc[shown["Conservé"] == "Non"]

    client_kept = int((view.loc[client_names == chosen, "Conservé"] == "Oui").sum())
    client_removed = int((view.loc[client_names == chosen, "Conservé"] == "Non").sum())
    kept_word = "conservée" if client_kept == 1 else "conservées"
    removed_word = "retirée" if client_removed == 1 else "retirées"
    st.write(f"{client_kept} {kept_word}, {client_removed} {removed_word}.")

    if shown.empty:
        st.info("Aucune ligne à afficher.")
    else:
        st.dataframe(shown, use_container_width=True, hide_index=True)


try:
    from streamlit.runtime.scriptrunner import get_script_run_ctx
    _in_streamlit = get_script_run_ctx() is not None
except Exception:
    _in_streamlit = False

if __name__ == "__main__" or _in_streamlit:
    main()
