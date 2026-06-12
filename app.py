from pathlib import Path
import tempfile
import random

import streamlit as st
import streamlit.components.v1 as components

from fill_hs_codes import fill_invoice, fill_pdf_invoice

st.set_page_config(page_title="HS-code invoice tool", layout="centered")

with st.sidebar:
    st.header("Instellingen")
    disco_active = st.toggle("Disco Mode 🪩")

if disco_active:
    st.markdown(
        """
        <style>
        @keyframes disco-bg {
            0%   { background-color: #ff0080; }
            14%  { background-color: #ff6600; }
            28%  { background-color: #ffee00; }
            42%  { background-color: #00cc44; }
            57%  { background-color: #00ccff; }
            71%  { background-color: #0055ff; }
            85%  { background-color: #aa00ff; }
            100% { background-color: #ff0080; }
        }
        @keyframes disco-text {
            0%   { color: #ff0080; }
            14%  { color: #ff6600; }
            28%  { color: #ffee00; }
            42%  { color: #00cc44; }
            57%  { color: #00ccff; }
            71%  { color: #0055ff; }
            85%  { color: #aa00ff; }
            100% { color: #ff0080; }
        }
        .stApp { animation: disco-bg 1.8s linear infinite; }
        .stApp h1, .stApp h2, .stApp h3 {
            animation: disco-text 1.2s linear infinite;
            font-weight: 900 !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

components.html(
    """
    <script>
    (function() {
        const p = window.parent;
        const pd = p.document;

        if (p.__easterEggInit) return;
        p.__easterEggInit = true;

        const gif = pd.createElement('img');
        gif.style.cssText = [
            'position:fixed', 'bottom:20px', 'left:20px', 'z-index:999999',
            'max-width:320px', 'max-height:320px', 'border-radius:12px',
            'box-shadow:0 8px 32px rgba(0,0,0,0.5)', 'display:none',
            'pointer-events:none', 'transition:opacity 0.6s',
        ].join(';');
        pd.body.appendChild(gif);

        let playing = false;
        let hideTimer = null;

        p.addEventListener('mousemove', function(e) {
            const inCorner = e.clientX < 80 && e.clientY > (p.innerHeight - 80);
            if (inCorner && !playing) {
                playing = true;
                gif.src = '/app/static/easter-egg.gif?t=' + Date.now();
                gif.style.display = 'block';
                gif.style.opacity = '1';
                clearTimeout(hideTimer);
                hideTimer = setTimeout(function() {
                    gif.style.opacity = '0';
                    setTimeout(function() {
                        gif.style.display = 'none';
                        gif.src = '';
                        playing = false;
                    }, 600);
                }, 4000);
            }
        });
    })();
    </script>
    """,
    height=0,
)

st.title("HS-code invoice tool")
st.write(
    "Upload een Excel- of PDF-invoice. Bij Excel behoudt de tool de originele invoice en zet direct naast "
    "**Article No.** een kolom **HS code** met de code uit de artikellijst. "
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

    stem = Path(invoice.name).stem
    is_final = "final" in stem.lower()
    if is_final:
        out_stem = stem + "_final_ECHT_DEFINITIEF_nu_echt"
    else:
        out_stem = stem + "_met_HS_codes"

    excel_spinners = [
        "Excel lezen en HS-codes invullen...",
        "Douane aan het bellen...",
        "HS-codes aan het verzinnen...",
        "Googlen wat dit artikel eigenlijk is...",
        "Undercover tolambtenaar raadplegen...",
        "Tariefboek van 1987 doorbladeren...",
        "Collega vragen die het ook niet weet...",
        "Gewoon een getal invullen en hopen...",
    ]
    pdf_spinners = [
        "PDF lezen en HS-codes zoeken...",
        "OCR aan het huilen maken...",
        "Pixels proberen te begrijpen...",
        "De scanner de schuld geven...",
        "Handschrift ontcijferen met tranen...",
        "Koffie halen terwijl OCR nadenkt...",
    ]

    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        in_path = tmp / invoice.name
        out_path = tmp / (out_stem + ".xlsx")
        map_path = tmp / "hs_mapping.csv"

        in_path.write_bytes(invoice.getvalue())
        if mapping:
            map_path.write_bytes(mapping.getvalue())
        else:
            map_path = Path(__file__).with_name("hs_mapping.csv")

        try:
            if in_path.suffix.lower() == ".pdf":
                with st.spinner(random.choice(pdf_spinners)):
                    result = fill_pdf_invoice(in_path, out_path, map_path)
            else:
                with st.spinner(random.choice(excel_spinners)):
                    result = fill_invoice(in_path, out_path, map_path)
            st.session_state["result"] = result
            st.session_state["output_name"] = out_path.name
            st.session_state["output_bytes"] = out_path.read_bytes()
            st.session_state["is_final"] = is_final
        except Exception as e:
            st.error(f"Er ging iets mis: {e}")

    result = st.session_state.get("result")
    output_bytes = st.session_state.get("output_bytes")
    output_name = st.session_state.get("output_name")

    if result and output_bytes and output_name:
        st.success(f"Klaar: {result['filled']} HS-codes ingevuld.")
        if st.session_state.get("is_final"):
            st.info("'Final' hè? We'll see about that. 😏")
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
