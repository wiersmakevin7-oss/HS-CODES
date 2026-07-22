from pathlib import Path
import tempfile

import streamlit as st

from fill_hs_codes import fill_invoice, fill_pdf_invoice

st.set_page_config(page_title="HS-code invoice tool", layout="centered")
st.title("HS-code invoice tool")
st.write(
    "Upload een Excel- of PDF-invoice. Bij Excel behoudt de tool de originele invoice en zet helemaal rechts "
    "een kolom **HS code (artikellijst)** met de code op basis van het artikelnummer. "
    "Bij PDF maakt de tool een nieuwe Excel met de gevonden artikelregels en HS-codes. "
    "Gescande PDF's worden met OCR gelezen en kunnen iets langer duren. "
    "HSN-nummers uit de invoice worden niet gebruikt."
)

invoice = st.file_uploader("Upload invoice (.xlsx of .pdf)", type=["xlsx", "pdf"])
mapping = st.file_uploader("Optioneel: upload nieuw artikeloverzicht / mapping CSV", type=["csv"])

if invoice:
    source_key = (invoice.name, invoice.size, mapping.name if mapping else "default")
    if st.session_state.get("source_key") != source_key:
        st.session_state.pop("output_bytes", None)
        st.session_state.pop("output_name", None)
        st.session_state.pop("result", None)
        st.session_state["source_key"] = source_key

    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        in_path = tmp / invoice.name
        out_path = tmp / (Path(invoice.name).stem + "_met_HS_codes.xlsx")
        map_path = tmp / "hs_mapping.csv"

        in_path.write_bytes(invoice.getvalue())
        if mapping:
            map_path.write_bytes(mapping.getvalue())
        else:
            map_path = Path(__file__).with_name("hs_mapping.csv")

        try:
            if in_path.suffix.lower() == ".pdf":
                with st.spinner("PDF lezen en HS-codes zoeken..."):
                    result = fill_pdf_invoice(in_path, out_path, map_path)
            else:
                with st.spinner("Excel lezen en HS-codes invullen..."):
                    result = fill_invoice(in_path, out_path, map_path)
            st.session_state["result"] = result
            st.session_state["output_name"] = out_path.name
            st.session_state["output_bytes"] = out_path.read_bytes()
        except Exception as e:
            st.error(f"Er ging iets mis: {e}")

    result = st.session_state.get("result")
    output_bytes = st.session_state.get("output_bytes")
    output_name = st.session_state.get("output_name")

    if result and output_bytes and output_name:
        message = f"Klaar: {result['filled']} HS-codes ingevuld."
        invoice_total = result.get("invoice_total")
        if isinstance(invoice_total, (int, float)):
            message += f" Factuurwaarde: {invoice_total:,.2f}."
        st.success(message)
        st.download_button(
            "Download nieuwe Excel met HS-codes",
            data=output_bytes,
            file_name=output_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
        if result["unmatched_count"]:
            st.warning(f"{result['unmatched_count']} artikelregels konden niet gekoppeld worden.")
            with st.expander("Bekijk niet-gekoppelde regels"):
                st.dataframe(result["unmatched"])
