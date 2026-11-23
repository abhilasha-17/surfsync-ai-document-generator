"""SurfSync Infotech Word document generator."""

import base64
import hashlib
import os
from io import BytesIO
from pathlib import Path

import streamlit as st
from dotenv import load_dotenv

from src.docx_filler import fill_docx
from src.generator import generate_content
from src.template_reader import inspect_docx


load_dotenv()

ROOT = Path(__file__).resolve().parent
LOGO_PATH = ROOT / "assets" / "surfsync_logo.png"

API_KEY = os.getenv("LLM_API_KEY", "").strip()
API_BASE_URL = os.getenv(
    "LLM_BASE_URL",
    "https://api.groq.com/openai/v1",
).strip()
MODEL = os.getenv("LLM_MODEL", "openai/gpt-oss-20b").strip()


st.set_page_config(
    page_title="SurfSync Infotech | Document Studio",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="collapsed",
)


if "dark_mode" not in st.session_state:
    st.session_state.dark_mode = False

if "generated_docx" not in st.session_state:
    st.session_state.generated_docx = None


switch_space, theme_column = st.columns([5, 1])

with theme_column:
    st.toggle("🌙 Dark mode", key="dark_mode")


dark = st.session_state.dark_mode

if dark:
    palette = {
        "page": "#0b1220",
        "glow": "rgba(19, 119, 145, .20)",
        "ink": "#edf5f7",
        "muted": "#a6bac4",
        "card": "rgba(19, 31, 48, .72)",
        "card_border": "rgba(164, 201, 211, .18)",
        "input": "rgba(8, 17, 29, .80)",
        "input_border": "rgba(164, 201, 211, .25)",
        "teal": "#22b8ca",
        "teal_deep": "#087f98",
        "soft": "rgba(28, 74, 89, .44)",
        "placeholder": "#8ea4af",
        "note_bg": "rgba(99, 73, 29, .25)",
        "note_border": "rgba(234, 190, 101, .30)",
        "note_text": "#f0d9a9",
        "logo_blend": "normal",
    }
else:
    palette = {
        "page": "#edf4f7",
        "glow": "rgba(16, 150, 173, .15)",
        "ink": "#17364a",
        "muted": "#607887",
        "card": "rgba(255, 255, 255, .72)",
        "card_border": "rgba(255, 255, 255, .90)",
        "input": "rgba(255, 255, 255, .80)",
        "input_border": "rgba(29, 76, 96, .17)",
        "teal": "#087f98",
        "teal_deep": "#145f78",
        "soft": "rgba(222, 241, 244, .78)",
        "placeholder": "#80939d",
        "note_bg": "rgba(255, 249, 237, .90)",
        "note_border": "rgba(210, 162, 75, .35)",
        "note_text": "#715727",
        "logo_blend": "multiply",
    }


CSS = """
<style>
  :root {
    color-scheme: __COLOR_SCHEME__;
    --page: __PAGE__;
    --glow: __GLOW__;
    --ink: __INK__;
    --muted: __MUTED__;
    --card: __CARD__;
    --card-border: __CARD_BORDER__;
    --input: __INPUT__;
    --input-border: __INPUT_BORDER__;
    --teal: __TEAL__;
    --teal-deep: __TEAL_DEEP__;
    --soft: __SOFT__;
    --placeholder: __PLACEHOLDER__;
    --note-bg: __NOTE_BG__;
    --note-border: __NOTE_BORDER__;
    --note-text: __NOTE_TEXT__;
    --logo-blend: __LOGO_BLEND__;
  }

  html,
  body,
  .stApp,
  [data-testid="stAppViewContainer"] {
    background:
      radial-gradient(ellipse at 5% 3%, var(--glow), transparent 36%),
      radial-gradient(ellipse at 97% 18%, rgba(237, 150, 60, .07), transparent 29%),
      var(--page) !important;
    color: var(--ink) !important;
  }

  [data-testid="stHeader"],
  [data-testid="stDecoration"],
  [data-testid="stToolbar"],
  [data-testid="stSidebar"] {
    display: none !important;
    height: 0 !important;
  }

  .block-container {
    width: 100%;
    max-width: none !important;
    padding: .35rem 4vw 2.25rem !important;
  }

  [data-testid="stToggle"] {
    padding: .52rem .82rem !important;
    margin: .15rem 0 .45rem auto !important;
    width: fit-content !important;
    min-width: 145px;
    border: 1px solid var(--card-border) !important;
    border-radius: 999px !important;
    background: var(--card) !important;
    box-shadow: 0 8px 24px rgba(20, 53, 69, .08) !important;
    backdrop-filter: blur(14px);
    -webkit-backdrop-filter: blur(14px);
  }

  [data-testid="stToggle"] label {
    display: flex !important;
    align-items: center !important;
    gap: .55rem !important;
    cursor: pointer !important;
  }

  [data-testid="stToggle"] label p {
    color: var(--ink) !important;
    font-size: .88rem !important;
    font-weight: 750 !important;
  }

  [data-testid="stToggle"] [role="switch"] {
    transform: scale(1.08);
  }

  [data-testid="stToggle"] [role="switch"][aria-checked="true"] {
    background: var(--teal) !important;
  }

  .hero {
    display: grid;
    grid-template-columns: minmax(150px, 230px) 1fr;
    align-items: center;
    gap: clamp(1.5rem, 4vw, 4rem);
    padding: .65rem 0 1.1rem;
    animation: fade-up .55s ease both;
  }

  .logo-frame {
    display: flex;
    align-items: center;
    justify-content: center;
    padding: .35rem;
    background: transparent;
    transform: rotate(-1deg);
  }

  .logo-frame img {
    display: block;
    width: 100%;
    max-width: 215px;
    height: auto;
    object-fit: contain;
    border-radius: 14px;
    mix-blend-mode: var(--logo-blend);
    filter: drop-shadow(0 12px 20px rgba(24, 54, 74, .16));
  }

  .brand-kicker {
    color: var(--teal);
    font-size: .76rem;
    font-weight: 850;
    letter-spacing: .20em;
    text-transform: uppercase;
    margin: 0 0 .65rem;
  }

  .hero-title {
    color: var(--ink);
    font-size: clamp(2.35rem, 4.5vw, 4.15rem);
    font-weight: 780;
    letter-spacing: -.055em;
    line-height: 1.03;
    margin: 0 0 .8rem;
  }

  .hero-copy {
    color: var(--muted);
    font-size: clamp(1rem, 1.35vw, 1.18rem);
    line-height: 1.65;
    max-width: 850px;
    margin: 0;
  }

  .hero-tag {
    display: inline-flex;
    margin-top: 1.2rem;
    padding: .52rem .9rem;
    border: 1px solid var(--card-border);
    border-radius: 999px;
    background: var(--card);
    color: var(--teal);
    font-size: .73rem;
    font-weight: 850;
    letter-spacing: .11em;
    backdrop-filter: blur(16px);
    -webkit-backdrop-filter: blur(16px);
  }

  .hero-bottom-space {
    height: 1.1rem;
  }

  .section-title {
    color: var(--ink) !important;
    font-size: 1.14rem;
    font-weight: 760;
    margin: 0 0 .38rem;
  }

  .section-copy {
    color: var(--muted) !important;
    font-size: .93rem;
    line-height: 1.6;
    margin: 0 0 .9rem;
  }

  .step-index {
    display: inline-flex;
    width: 1.78rem;
    height: 1.78rem;
    margin-right: .55rem;
    align-items: center;
    justify-content: center;
    border-radius: 50%;
    background: var(--soft);
    color: var(--teal);
    font-size: .77rem;
    font-weight: 850;
    vertical-align: middle;
  }

  div[data-testid="stVerticalBlockBorderWrapper"] {
    min-height: 430px !important;
    box-sizing: border-box !important;
    background: var(--card) !important;
    border: 1px solid var(--card-border) !important;
    border-radius: 22px !important;
    box-shadow: 0 18px 55px rgba(22, 48, 64, .10) !important;
    backdrop-filter: blur(22px);
    -webkit-backdrop-filter: blur(22px);
    animation: fade-up .55s ease both;
  }

  div[data-testid="stVerticalBlockBorderWrapper"]:hover {
    box-shadow: 0 22px 64px rgba(22, 48, 64, .14) !important;
  }

  div[data-testid="stFileUploader"] section {
    background: var(--input) !important;
    border: 1px dashed var(--input-border) !important;
    border-radius: 15px !important;
  }

  div[data-testid="stFileUploader"] button {
    background: linear-gradient(110deg, var(--teal-deep), #174e67) !important;
    border: 0 !important;
    border-radius: 10px !important;
    color: #fff !important;
    font-weight: 700 !important;
  }

  div[data-testid="stFileUploader"] small,
  div[data-testid="stFileUploader"] span {
    color: var(--muted) !important;
  }

  [data-baseweb="select"] > div {
    background: var(--input) !important;
    color: var(--ink) !important;
    border-color: var(--input-border) !important;
    border-radius: 12px !important;
    min-height: 3rem;
  }

  .stTextArea textarea {
    min-height: 260px !important;
    background: var(--input) !important;
    color: var(--ink) !important;
    border: 1px solid var(--input-border) !important;
    border-radius: 14px !important;
    line-height: 1.65 !important;
    backdrop-filter: blur(12px);
  }

  .stTextArea textarea::placeholder {
    color: var(--placeholder) !important;
  }

  div.stButton > button[kind="primary"] {
    min-height: 3.4rem;
    border: 1px solid rgba(255, 255, 255, .22) !important;
    border-radius: 13px !important;
    background: linear-gradient(105deg, var(--teal-deep), #164e69) !important;
    color: #fff !important;
    font-size: 1rem !important;
    font-weight: 760 !important;
    letter-spacing: .015em;
    box-shadow: 0 12px 28px rgba(8, 127, 152, .23) !important;
    transition: transform .2s ease, box-shadow .2s ease, filter .2s ease;
  }

  div.stButton > button[kind="primary"]:hover {
    filter: brightness(1.08);
    transform: translateY(-2px);
    box-shadow: 0 16px 34px rgba(8, 127, 152, .31) !important;
  }

  div.stButton > button[kind="primary"]:disabled {
    background: #84969d !important;
    box-shadow: none !important;
    color: #fff !important;
    transform: none;
  }

  div[data-testid="stDownloadButton"] > button {
    width: 100%;
    min-height: 3.15rem;
    border: 1px solid rgba(255, 255, 255, .22) !important;
    border-radius: 12px !important;
    background: linear-gradient(110deg, var(--teal-deep), #164e69) !important;
    color: #fff !important;
    font-weight: 760 !important;
    box-shadow: 0 9px 22px rgba(8, 127, 152, .20) !important;
    transition: transform .2s ease, box-shadow .2s ease;
  }

  div[data-testid="stDownloadButton"] > button p,
  div[data-testid="stDownloadButton"] > button span,
  div[data-testid="stDownloadButton"] > button div {
    color: #fff !important;
  }

  div[data-testid="stDownloadButton"] > button:hover {
    background: linear-gradient(110deg, #096f84, #103f56) !important;
    color: #fff !important;
    transform: translateY(-2px);
    box-shadow: 0 13px 28px rgba(8, 127, 152, .28) !important;
  }

  .service-note {
    padding: .85rem 1rem;
    border-radius: 13px;
    border: 1px solid var(--note-border);
    background: var(--note-bg);
    color: var(--note-text);
    font-size: .92rem;
    line-height: 1.5;
    margin: .7rem 0 1.1rem;
  }

  div[data-testid="stAlert"] {
    border-radius: 13px;
  }

  .footer-copy {
    color: var(--muted);
    font-size: .83rem;
    text-align: center;
    margin-top: 1rem;
  }

  footer {
    visibility: hidden;
  }

  @keyframes fade-up {
    from { opacity: 0; transform: translateY(9px); }
    to { opacity: 1; transform: translateY(0); }
  }

  @media (prefers-reduced-motion: reduce) {
    *,
    *::before,
    *::after {
      animation-duration: .01ms !important;
      animation-iteration-count: 1 !important;
      transition-duration: .01ms !important;
    }
  }

  @media (max-width: 850px) {
    .block-container {
      padding: .35rem 5vw 2rem !important;
    }

    .hero {
      grid-template-columns: 145px 1fr;
      gap: 1.3rem;
    }

    .hero-title {
      font-size: 2.5rem;
    }

    div[data-testid="stVerticalBlockBorderWrapper"] {
      min-height: 0 !important;
    }
  }

  @media (max-width: 600px) {
    .block-container {
      padding: .25rem 1rem 1.5rem !important;
    }

    .hero {
      grid-template-columns: 1fr;
      gap: .75rem;
    }

    .logo-frame {
      justify-content: flex-start;
    }

    .logo-frame img {
      max-width: 175px;
    }

    .hero-title {
      font-size: 2.05rem;
    }

    [data-testid="stToggle"] {
      margin-left: 0 !important;
    }
  }
</style>
"""

replacements = {
    "__COLOR_SCHEME__": "dark" if dark else "light",
    "__PAGE__": palette["page"],
    "__GLOW__": palette["glow"],
    "__INK__": palette["ink"],
    "__MUTED__": palette["muted"],
    "__CARD__": palette["card"],
    "__CARD_BORDER__": palette["card_border"],
    "__INPUT__": palette["input"],
    "__INPUT_BORDER__": palette["input_border"],
    "__TEAL__": palette["teal"],
    "__TEAL_DEEP__": palette["teal_deep"],
    "__SOFT__": palette["soft"],
    "__PLACEHOLDER__": palette["placeholder"],
    "__NOTE_BG__": palette["note_bg"],
    "__NOTE_BORDER__": palette["note_border"],
    "__NOTE_TEXT__": palette["note_text"],
    "__LOGO_BLEND__": palette["logo_blend"],
}

for token, value in replacements.items():
    CSS = CSS.replace(token, value)

st.markdown(CSS, unsafe_allow_html=True)


# Display the logo and hero together so no empty wrapper appears above the logo.
logo_html = ""

if LOGO_PATH.is_file():
    logo_base64 = base64.b64encode(LOGO_PATH.read_bytes()).decode("ascii")
    logo_html = (
        '<div class="logo-frame">'
        f'<img src="data:image/png;base64,{logo_base64}" '
        'alt="SurfSync Infotech logo">'
        "</div>"
    )

st.markdown(
    f"""
    <section class="hero">
      {logo_html}
      <div class="hero-content">
        <div class="brand-kicker">SurfSync Infotech</div>
        <h1 class="hero-title">Client Document Studio</h1>
        <p class="hero-copy">
          Turn a project brief into a polished client-ready Word document,
          using your existing template.
        </p>
        <span class="hero-tag">
          BRD&nbsp;&nbsp;·&nbsp;&nbsp;FRD&nbsp;&nbsp;·&nbsp;&nbsp;SRS
          &nbsp;&nbsp;·&nbsp;&nbsp;QUOTATION
        </span>
      </div>
    </section>
    <div class="hero-bottom-space"></div>
    """,
    unsafe_allow_html=True,
)


if not API_KEY:
    st.markdown(
        """
        <div class="service-note">
          Document generation is temporarily unavailable.
          Please try again later.
        </div>
        """,
        unsafe_allow_html=True,
    )


left, right = st.columns([1, 1], gap="large")

uploaded = None
document_type = "BRD"
project_prompt = ""
template_info = None
template_bytes = None

with left:
    with st.container(border=True):
        st.markdown(
            """
            <div class="section-title">
              <span class="step-index">01</span>Choose your template
            </div>
            <div class="section-copy">
              Upload a branded Word document. Your original file is never overwritten.
            </div>
            """,
            unsafe_allow_html=True,
        )

        uploaded = st.file_uploader(
            "Word template",
            type=["docx"],
            label_visibility="collapsed",
        )

        st.markdown('<div style="height:1.25rem"></div>', unsafe_allow_html=True)

        st.markdown(
            """
            <div class="section-title">
              <span class="step-index">02</span>Select document type
            </div>
            <div class="section-copy">
              Choose the format that matches your template.
            </div>
            """,
            unsafe_allow_html=True,
        )

        document_type = st.selectbox(
            "Document type",
            ["BRD", "FRD", "SRS", "Quotation"],
            label_visibility="collapsed",
        )

    if uploaded is not None:
        template_bytes = uploaded.getvalue()

        try:
            template_info = inspect_docx(BytesIO(template_bytes))

            st.success(
                f"Template ready · {len(template_info['placeholders'])} placeholder(s) · "
                f"{len(template_info['tables'])} table(s)"
            )

            with st.expander("Review detected template structure"):
                st.write(
                    "**Headings**",
                    template_info["headings"] or "No Word heading styles detected",
                )
                st.write(
                    "**Placeholders**",
                    template_info["placeholders"] or "None detected",
                )

                for table_info in template_info["tables"]:
                    st.write(
                        f"**Table {table_info['index']} columns**",
                        table_info["columns"],
                    )

                if not template_info["placeholders"]:
                    st.warning(
                        "For reliable placement, use template placeholders such as "
                        "{{client_name}} and {{scope_of_work}}."
                    )

        except Exception as exc:
            template_bytes = None
            template_info = None
            st.error(f"This file could not be read as a valid Word document: {exc}")


with right:
    with st.container(border=True):
        st.markdown(
            """
            <div class="section-title">
              <span class="step-index">03</span>Describe the project
            </div>
            <div class="section-copy">
              Include the client, goals, scope, constraints, timeline,
              deliverables, and confirmed pricing where applicable.
            </div>
            """,
            unsafe_allow_html=True,
        )

        project_prompt = st.text_area(
            "Project details",
            height=260,
            label_visibility="collapsed",
            placeholder=(
                "Describe the client, goals, required features, users, constraints, "
                "timeline, deliverables, and any confirmed pricing."
            ),
        )

        st.markdown(
            """
            <div class="section-copy">
              Include only confirmed prices and dates. Review the document before sharing.
            </div>
            """,
            unsafe_allow_html=True,
        )


template_bytes_for_signature = template_bytes or b""
signature_payload = (
    template_bytes_for_signature
    + b"\\0"
    + document_type.encode("utf-8")
    + b"\\0"
    + project_prompt.encode("utf-8")
)
current_signature = hashlib.sha256(signature_payload).hexdigest()

st.markdown('<div style="height:.9rem"></div>', unsafe_allow_html=True)

can_generate = bool(
    uploaded
    and project_prompt.strip()
    and template_info
    and API_KEY
)

generate_clicked = st.button(
    "Generate Word document",
    type="primary",
    use_container_width=True,
    disabled=not can_generate,
)


if generate_clicked:
    st.session_state.generated_docx = None

    try:
        with st.status("Preparing your Word document…", expanded=True) as status:
            st.write("Structuring the brief for the selected document type.")

            generated = generate_content(
                base_url=API_BASE_URL,
                model=MODEL,
                api_key=API_KEY,
                document_type=document_type,
                project_prompt=project_prompt,
                template_info=template_info,
            )

            st.write("Filling a copy of your Word template.")

            docx_result = fill_docx(
                template_bytes,
                generated.fields,
                generated.items,
            )

            status.update(
                label="Your Word document is ready",
                state="complete",
                expanded=False,
            )

        output_stem = f"surfsync_{document_type.lower()}_generated"

        st.session_state.generated_docx = {
            "signature": current_signature,
            "data": docx_result,
            "file_name": f"{output_stem}.docx",
        }

    except Exception as exc:
        st.error(f"Generation failed: {exc}")


saved_document = st.session_state.generated_docx

if saved_document and saved_document["signature"] == current_signature:
    st.markdown("### Your Word document is ready")

    st.download_button(
        "Download Word document",
        data=saved_document["data"],
        file_name=saved_document["file_name"],
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "wordprocessingml.document"
        ),
        use_container_width=True,
    )

elif saved_document and saved_document["signature"] != current_signature:
    st.info("Your project details changed. Generate a new document to download the updated version.")


st.markdown(
    """
    <p class="footer-copy">
      SurfSync Infotech · Review requirements, assumptions, dates, prices,
      and contractual terms before sharing.
    </p>
    """,
    unsafe_allow_html=True,
)