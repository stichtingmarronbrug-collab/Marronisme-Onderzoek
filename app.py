
import os
from datetime import datetime
import pandas as pd
import streamlit as st

TEMPLATE_XLSX = "marronisme_onderzoek_template.xlsx"

DATA_DIR = "data"
os.makedirs(DATA_DIR, exist_ok=True)
OUT_CSV = os.path.join(DATA_DIR, "responses_formulier.csv")

DIM_ORDER = ["historisch", "sociaal", "moreel", "spiritueel", "praktisch"]

def load_template():
    q = pd.read_excel(TEMPLATE_XLSX, sheet_name="questions", dtype=str).fillna("")
    m = pd.read_excel(TEMPLATE_XLSX, sheet_name="mc_options", dtype=str).fillna("")

    q["vraag_id"] = q["vraag_id"].str.strip()
    q["dimensie"] = q["dimensie"].str.strip().str.lower()
    q["type"] = q["type"].str.strip().str.lower()

    # mc opties
    m["vraag_id"] = m["vraag_id"].str.strip()
    m["optie"] = m["optie"].str.strip().str.upper()
    m["label"] = m["label"].str.strip()

    mc_map = {}
    for vid, sub in m.groupby("vraag_id"):
        # sorteer A/B/C/D
        order = {"A": 1, "B": 2, "C": 3, "D": 4}
        sub = sub.copy()
        sub["rank"] = sub["optie"].map(order).fillna(99)
        sub = sub.sort_values("rank")

        mc_map[vid] = [f"{row.optie} — {row.label}" for row in sub.itertuples(index=False)]

    # sorteer vragen Q01..Q100 binnen dimensies
    def qnum(v):
        try:
            return int(v.replace("Q", "").strip())
        except Exception:
            return 9999

    q["qnum"] = q["vraag_id"].apply(qnum)
    dim_rank = {d: i for i, d in enumerate(DIM_ORDER)}
    q["dim_rank"] = q["dimensie"].map(dim_rank).fillna(999)
    q = q.sort_values(["dim_rank", "qnum"]).drop(columns=["qnum", "dim_rank"])

    return q, mc_map

def append_rows(rows):
    df = pd.DataFrame(rows)[["respondent_id","vraag_id","dimensie","type","antwoord","timestamp"]]
    try:
        if os.path.exists(OUT_CSV):
            df.to_csv(OUT_CSV, mode="a", header=False, index=False, encoding="utf-8-sig")
        else:
            df.to_csv(OUT_CSV, index=False, encoding="utf-8-sig")
    except PermissionError:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        alt = os.path.join(DATA_DIR, f"responses_formulier_{ts}.csv")
        df.to_csv(alt, index=False, encoding="utf-8-sig")

def main():
    st.set_page_config(page_title="Marronisme Onderzoek", layout="wide")
    st.title("Marronisme Onderzoek – Enquête (alle vragen)")
    st.caption("Alle vragen worden geladen uit het Excel-template (questions + mc_options).")

    if not os.path.exists(TEMPLATE_XLSX):
        st.error(f"Template niet gevonden: {TEMPLATE_XLSX}")
        st.stop()

    questions, mc_map = load_template()
    st.write("Totaal vragen geladen:", len(questions))

    respondent_id = st.text_input("Respondent ID (bijv. 001)")

    with st.form("form_all"):
        answers = []

        for dim in DIM_ORDER:
            sub = questions[questions["dimensie"] == dim]
            with st.expander(f"{dim.capitalize()} ({len(sub)} vragen)", expanded=(dim=="historisch")):
                for _, q in sub.iterrows():
                    vid = q["vraag_id"]
                    qtype = q["type"]
                    vtext = str(q.get("vraagtekst", "")).strip()

                    st.markdown(f"### {vid}")
                    if vtext:
                        st.write(vtext)

                    if qtype == "mc":
                        options = mc_map.get(vid, ["A", "B", "C", "D"])
                        choice = st.radio("Kies één:", options, key=f"ans_{vid}")
                        answer = str(choice).split("—")[0].strip()
                    else:
                        answer = st.text_area("Antwoord:", key=f"ans_{vid}", height=120)

                    answers.append((vid, dim, qtype, answer))
                    st.divider()

        submitted = st.form_submit_button("Alles opslaan", type="primary")

        if submitted:
            if not respondent_id.strip():
                st.error("Respondent ID is verplicht.")
                st.stop()

            ts = datetime.now().isoformat(timespec="seconds")
            rows = []
            for vid, dim, t, ans in answers:
                rows.append({
                    "respondent_id": respondent_id.strip(),
                    "vraag_id": vid,
                    "dimensie": dim,
                    "type": t,
                    "antwoord": (ans or "").strip(),
                    "timestamp": ts
                })

            append_rows(rows)
            st.success(f"Opgeslagen! → {OUT_CSV}")

if __name__ == "__main__":
    main()
