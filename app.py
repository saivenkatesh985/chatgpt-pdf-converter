import datetime
import os
import re
import tempfile
import markdown
import requests
import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from playwright.sync_api import sync_playwright

# Optional library imports with graceful fallbacks
try:
    from deep_translator import GoogleTranslator
    HAS_TRANSLATOR = True
except ImportError:
    HAS_TRANSLATOR = False

try:
    from gtts import gTTS
    HAS_GTTS = True
except ImportError:
    HAS_GTTS = False

# --- Automatically install playwright chromium browser on cloud deployment ---
@st.cache_resource
def install_playwright():
    try:
        import subprocess
        import sys
        subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"], check=True)
    except Exception as e:
        print(f"Playwright browser install notice: {e}")

install_playwright()

# Page config
st.set_page_config(
    page_title="ChatGPT Document Studio Pro", page_icon="⬇️", layout="wide"
)

# Custom UI polish
st.markdown(
    """
    <style>
        .stButton>button {
            border-radius: 8px;
            font-weight: 600;
            transition: all 0.2s ease-in-out;
        }
        .stButton>button:hover {
            transform: translateY(-1px);
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
        }
        .block-container {
            padding-top: 2rem;
        }
    </style>
""",
    unsafe_allow_html=True,
)

st.title("⬇️ ChatGPT Link to Multi-Format Document Studio Pro")
st.markdown(
    "Transform single or multiple public ChatGPT conversations into professionally styled documents "
    "with robust image rendering, translation, Mermaid vector diagrams, text-to-speech, and study revision toggles."
)

# --- SIDEBAR: ORGANIZED LAYOUT STUDIO ---
st.sidebar.header("🎨 Layout & Studio Controls")

with st.sidebar.expander("📌 Document & Branding", expanded=True):
    theme = st.selectbox(
        "Color Theme",
        ["White", "Black (OLED)", "Notes", "Study", "Conversation"],
    )
    layout_view_style = st.selectbox(
        "Layout Presentation",
        ["Professional Document / Essay", "Modern Chat Bubbles"],
    )
    font_style = st.selectbox("Font Style", ["Sans-serif", "Serif", "Monospace"])
    custom_title = st.text_input(
        "Document Title Text", value=""
    )
    header_alignment = st.selectbox(
        "Header Alignment", ["Left", "Center", "Right"]
    )
    footer_text = st.text_input("Custom Footer Text", value="")
    watermark_text = st.text_input("Center Page Watermark Text", value="")
    watermark_opacity = st.slider(
        "Watermark Opacity (%)", min_value=1, max_value=15, value=4, step=1
    )
    
    include_footer_wm = st.checkbox("Include Bottom Footer Watermark", value=False)
    footer_watermark_text = st.text_input(
        "Bottom Footer Watermark Text", 
        value="Exported by <a href='https://chatgpt-pdf-converter.streamlit.app' target='_blank' style='color: inherit; text-decoration: underline;'>ChatGPT Studio Pro</a>"
    )

with st.sidebar.expander("🌐 Advanced Features & Toggles", expanded=True):
    # Multi-Language Translation
    enable_translation = st.checkbox("Enable Live Translation", value=False)
    target_language = st.selectbox(
        "Target Language",
        ["Spanish", "French", "German", "Japanese", "Chinese (Simplified)", "Italian", "Portuguese", "Russian"],
        disabled=not enable_translation
    )
    lang_code_map = {
        "Spanish": "es", "French": "fr", "German": "de", "Japanese": "ja",
        "Chinese (Simplified)": "zh-CN", "Italian": "it", "Portuguese": "pt", "Russian": "ru"
    }

    st.markdown("---")
    # Clean On/Off Toggles for Study Revision & Custom Notes
    enable_study_revision = st.toggle(
        "📚 Study Notes & Flashcards Auto-Append", 
        value=False,
        help="Scans chat content for definitions and key tips to build a summary sheet at the end."
    )
    
    enable_custom_notes = st.toggle(
        "📝 Custom Note Injector", 
        value=False,
        help="Enables the manager to insert personal markdown notes between conversation turns."
    )

with st.sidebar.expander("📐 Spacing & Geometry", expanded=False):
    font_size = st.text_input("Font Size", value="16px")
    margin_mm = st.number_input(
        "Page Margins (mm)", min_value=0, max_value=50, value=15, step=1
    )
    max_content_width = st.slider(
        "Max Content Width (px)", min_value=600, max_value=1400, value=900, step=50
    )
    line_height = st.slider(
        "Line Spacing", min_value=1.2, max_value=2.0, value=1.6, step=0.1
    )
    paragraph_spacing = st.slider(
        "Paragraph Spacing (px)", min_value=4, max_value=24, value=12, step=2
    )

with st.sidebar.expander("✨ Styling, Code & Accents", expanded=False):
    accent_color_map = {
        "Modern Green (#10a37f)": "#10a37f",
        "Executive Blue (#2563eb)": "#2563eb",
        "Royal Purple (#7c3aed)": "#7c3aed",
        "Sunset Orange (#ea580c)": "#ea580c",
    }
    chosen_accent_label = st.selectbox(
        "Accent & Border Color", list(accent_color_map.keys())
    )
    chosen_accent = accent_color_map[chosen_accent_label]
    code_theme = st.selectbox(
        "Code Block Theme",
        ["GitHub", "Atom One Dark", "Monokai Sublime", "VS Code Dark"],
    )
    block_radius = st.selectbox(
        "Message Box Corners", ["Sharp (0px)", "Rounded (6px)", "Pill (12px)"]
    )
    enable_box_shadow = st.checkbox("Add Subtle Shadow to Boxes", value=True)
    custom_css_injection = st.text_area(
        "Custom CSS Injector (Optional)",
        placeholder="e.g., body { filter: contrast(110%); }",
        height=70,
    )

with st.sidebar.expander("⚙️ Structure & Filters", expanded=False):
    export_format = st.selectbox(
        "Export Format Type",
        [
            "PDF Document (.pdf)",
            "Word Document (.docx)",
            "Markdown Document (.md)",
            "HTML Webpage (.html)",
        ],
    )

    message_filter_role = st.selectbox(
        "Message Role Filter", ["All", "User", "Assistant"]
    )

    message_selection_mode = st.radio(
        "Inclusion Scope",
        ["All Messages", "Pick Specific Messages (Custom)"],
        horizontal=True,
    )

    enable_ai_filter = st.checkbox(
        "Enable Custom AI Closing Filter", value=True
    )
    generate_toc = st.checkbox("Include Table of Contents", value=True)
    show_page_numbers = st.checkbox("Include Page Numbers", value=True)
    show_header_title = st.checkbox("Include Title Header", value=True)
    show_dividers = st.checkbox("Show Turn Dividers", value=False)
    show_timestamp = st.checkbox("Include Export Timestamp", value=True)

# --- SESSION STATE ---
if "chat_sections" not in st.session_state:
    st.session_state.chat_sections = None

if "bulk_links_data" not in st.session_state:
    st.session_state.bulk_links_data = []

if "injected_notes" not in st.session_state:
    st.session_state.injected_notes = {}


# --- HELPER FUNCTIONS ---
def clean_assistant_text(text: str) -> str:
    text = re.sub(r".*?", "", text)
    pattern = re.compile(
        r"\n+(?:If you want, I can:|Would you like me to:|Let me know if you).*$",
        re.DOTALL | re.IGNORECASE,
    )
    return pattern.sub("", text).strip()


def translate_text(text: str, dest_lang_code: str) -> str:
    if not HAS_TRANSLATOR or not text.strip():
        return text
    try:
        translator = GoogleTranslator(source='auto', target=dest_lang_code)
        if len(text) > 4000:
            chunks = [text[i:i+4000] for i in range(0, len(text), 4000)]
            translated_chunks = [translator.translate(c) for c in chunks]
            return "".join(translated_chunks)
        else:
            return translator.translate(text)
    except Exception as e:
        print(f"Translation error: {e}")
        return text


def parse_math_expression_blocks(text: str) -> str:
    text = text.replace(r'\tilde{[f}', 'f').replace(r'\tilde{f}', 'f')
    text = text.replace(r'\rightarrow', ' → ').replace(r'\to', ' → ')
    text = re.sub(r'\\text\{([^}]+)\}', r'\1', text)

    def boxed_replacer(match):
        formula = match.group(1).replace(r'\text{', '').replace('}', '')
        return f'''
        <div style="font-family: Consolas, Monaco, monospace; background: rgba(37, 99, 235, 0.08); 
                    border: 1px solid #2563eb; padding: 8px 14px; margin: 10px 0; 
                    border-radius: 6px; font-weight: bold; text-align: center; font-size: 1.05em;">
            {formula}
        </div>
        '''
    text = re.sub(r'\[\s*\\boxed\{([^}]+)\}\s*\]', boxed_replacer, text)
    text = re.sub(r'\\boxed\{([^}]+)\}', r'\1', text)
    return text


def parse_custom_tags_and_flowcharts(text: str) -> str:
    text = parse_math_expression_blocks(text)
    text = re.sub(r'<icon[^>]*/>', '<div style="text-align: center; font-weight: bold; margin: 4px 0;">⬇️</div>', text)
    text = re.sub(r'</?box[^>]*>', '', text)
    text = re.sub(r'<text[^>]*>(.*?)</text>', r'<div style="margin: 4px 0;">\1</div>', text, flags=re.DOTALL | re.IGNORECASE)
    return text


def scrape_single_chat(url: str, apply_filter: bool, target_lang: str = None) -> list:
    clean_url = url.split("?")[0].strip()
    if not clean_url:
        return []

    parts = clean_url.rstrip("/").split("/")
    share_id = parts[-1]

    if not share_id or len(share_id) < 10:
        raise Exception(f"Invalid share link format: {clean_url}")

    api_endpoint = f"https://chatgpt.com/backend-api/share/{share_id}"
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML,"
            " like Gecko) Chrome/124.0.0.0 Safari/537.36"
        ),
        "Accept": "application/json",
    }

    try:
        response = requests.get(api_endpoint, headers=headers, timeout=15)
        if response.status_code == 404:
            raise Exception("Share link not found or expired.")
        if response.status_code != 200:
            raise Exception(f"API returned status code {response.status_code}")

        data = response.json()
    except Exception as e:
        raise Exception(f"Failed to fetch conversation data: {e}")

    messages = []
    try:
        mapping = data.get("mapping", {})
        for node_id, node_data in mapping.items():
            message = node_data.get("message")
            if message:
                author = message.get("author", {}).get("role")
                if author in ["user", "assistant"]:
                    content = message.get("content", {})
                    content_parts = content.get("parts", [])
                    extracted_texts = []
                    
                    metadata = message.get("metadata", {})
                    dalle_data = metadata.get("dalle", {})
                    if dalle_data and dalle_data.get("gen_image_url"):
                        dalle_url = dalle_data.get("gen_image_url")
                        extracted_texts.append(f"\n\n![DALL-E Image]({dalle_url})\n\n")

                    for p in content_parts:
                        if isinstance(p, str):
                            extracted_texts.append(p)
                        elif isinstance(p, dict):
                            img_url = (
                                p.get("url") 
                                or p.get("file_url") 
                                or p.get("download_url")
                                or metadata.get("image_url")
                            )
                            asset_ptr = p.get("asset_pointer")
                            if not img_url and asset_ptr:
                                if asset_ptr.startswith("http://") or asset_ptr.startswith("https://"):
                                    img_url = asset_ptr

                            if img_url:
                                extracted_texts.append(f"\n\n![Image/Attachment]({img_url})\n\n")
                            elif "text" in p:
                                extracted_texts.append(str(p["text"]))
                                
                    text_content = "".join(extracted_texts)
                    text_content = re.sub(r".*?", "", text_content)

                    # --- CLEANING FILTERS ---
                    text_content = re.sub(r"!\[.*?\]\(.*?\)", "", text_content)
                    text_content = re.sub(
                        r"(?:the\s+output\s+of\s+this\s+plugin\s+was\s+redacted\.?|reference:\s*operating\s+system\s+concepts[^.\n]*\.)",
                        "",
                        text_content,
                        flags=re.IGNORECASE
                    )
                    text_content = text_content.strip()
                    # ------------------------

                    if text_content.strip():
                        role_label = "User" if author == "user" else "ChatGPT"
                        if role_label == "ChatGPT" and apply_filter:
                            text_content = clean_assistant_text(text_content)

                        if target_lang and HAS_TRANSLATOR:
                            text_content = translate_text(text_content, target_lang)

                        if text_content.strip():
                            parsed_text = parse_custom_tags_and_flowcharts(text_content)
                            formatted_html = markdown.markdown(
                                parsed_text, extensions=["fenced_code", "tables"]
                            )
                            msg_id = f"{role_label}: {text_content[:60]}..."
                            messages.append({
                                "id": msg_id,
                                "role": role_label,
                                "text": text_content,
                                "html": formatted_html,
                            })
    except Exception as parse_err:
        raise Exception(f"Error parsing conversation JSON: {parse_err}")

    if not messages:
        raise Exception("No messages found in this share link.")
    return messages


def add_docx_page_number(run):
    fldChar1 = OxmlElement("w:fldChar")
    fldChar1.set(qn("w:fldCharType"), "begin")
    instrText = OxmlElement("w:instrText")
    instrText.set(qn("xml:space"), "preserve")
    instrText.text = "PAGE"
    fldChar2 = OxmlElement("w:fldChar")
    fldChar2.set(qn("w:fldCharType"), "separate")
    fldChar3 = OxmlElement("w:fldChar")
    fldChar3.set(qn("w:fldCharType"), "end")

    r = run._r
    r.append(fldChar1)
    r.append(instrText)
    r.append(fldChar2)
    r.append(fldChar3)


def get_filtered_sections(
    chat_sections, message_filter_role, selection_mode, allowed_ids
):
    if not chat_sections:
        return []

    filtered_sections = []
    for sec_title, messages in chat_sections:
        new_messages = []
        for msg in messages:
            role = msg["role"]
            if message_filter_role == "User" and role != "User":
                continue
            if message_filter_role == "Assistant" and role != "ChatGPT":
                continue

            if selection_mode == "Pick Specific Messages (Custom)":
                if msg.get("id") not in allowed_ids:
                    continue

            new_messages.append(msg)
        if new_messages:
            filtered_sections.append((sec_title, new_messages))
    return filtered_sections


def generate_study_revision_sheet(chat_sections) -> str:
    revision_items = []
    for _, messages in chat_sections:
        for msg in messages:
            if msg["role"] == "ChatGPT":
                lines = msg["text"].split("\n")
                for line in lines:
                    if line.strip().startswith(("* ", "- ", "1. ", "• ")) or "define" in line.lower() or "key" in line.lower() or "important" in line.lower():
                        if len(line.strip()) > 10 and line.strip() not in revision_items:
                            revision_items.append(line.strip())

    if not revision_items:
        revision_items = ["No key definitions or flashcard items automatically detected. General review of conversation recommended."]

    items_html = "".join([f"<li>{item}</li>" for item in revision_items[:25]])
    return f"""
    <div style="background: rgba(16, 163, 127, 0.05); border: 2px solid #10a37f; padding: 20px; border-radius: 8px; margin-top: 40px;">
        <h3 style="color: #10a37f; margin-top: 0;">📚 Quick Revision Sheet & Key Takeaways</h3>
        <p><i>Auto-generated study flashcard notes extracted from conversation payload:</i></p>
        <ul style="margin-bottom: 0;">
            {items_html}
        </ul>
    </div>
    """


def build_pdf_document(full_html: str, margin_val: int) -> bytes:
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            args=[
                "--disable-gpu",
                "--no-sandbox",
                "--disable-setuid-sandbox",
                "--disable-dev-shm-usage"
            ]
        )
        page = browser.new_page()
        page.set_content(full_html, wait_until="load")
        
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            temp_path = tmp.name

        m_str = f"{margin_val}mm"
        pdf_margins = {"top": m_str, "bottom": m_str, "left": m_str, "right": m_str}

        page.pdf(path=temp_path, format="Letter", margin=pdf_margins, print_background=True)
        browser.close()

    with open(temp_path, "rb") as f:
        pdf_bytes = f.read()
    if os.path.exists(temp_path):
        os.remove(temp_path)
    return pdf_bytes


def build_html_document(
    chat_sections: list,
    theme: str,
    layout_view_style: str,
    font_style: str,
    font_size: str,
    line_height: float,
    paragraph_spacing: int,
    message_filter_role: str,
    selection_mode: str,
    allowed_ids: list,
    show_header_title: bool,
    custom_title: str,
    header_alignment: str,
    show_dividers: bool,
    show_timestamp: bool,
    footer_text: str,
    chosen_accent: str,
    code_theme_name: str,
    block_radius: str,
    margin_val: int,
    max_width: int,
    shadow_enabled: bool,
    include_toc: bool,
    watermark: str,
    watermark_alpha: float,
    footer_wm: str,
    include_page_nos: bool,
    custom_css: str,
    include_study_sheet: bool,
    injected_notes_dict: dict,
) -> str:
    active_sections = get_filtered_sections(
        chat_sections, message_filter_role, selection_mode, allowed_ids
    )

    if theme == "Black (OLED)":
        bg_color, text_color = "#000000", "#f3f4f6"
        user_bg, user_text = "#121212", "#e4e4e7"
        border_color, code_bg, code_color = "#27272a", "#121212", "#f3f4f6"
    elif theme == "Notes":
        bg_color, text_color = "#fbf0d9", "#2c221e"
        user_bg, user_text = "#e6f2ed", "#1b4d3e"
        border_color, code_bg, code_color = "#e6d7bc", "#f4e7ce", "#2c221e"
    elif theme == "Study":
        bg_color, text_color = "#f0f4f8", "#102a43"
        user_bg, user_text = "#d9e2ec", "#243b53"
        border_color, code_bg, code_color = "#bcccdc", "#e2e8f0", "#102a43"
    elif theme == "Conversation":
        bg_color, text_color = "#f7f7f8", "#202123"
        user_bg, user_text = "#eef8f6", "#0f5132"
        border_color, code_bg, code_color = "#e5e5e6", "#ffffff", "#202123"
    else:  # White
        bg_color, text_color = "#ffffff", "#1a1a1a"
        user_bg, user_text = "#edf7f0", "#114f2e"
        border_color, code_bg, code_color = "#e1e4e8", "#f6f8fa", "#24292e"

    code_theme_map = {
        "GitHub": "github",
        "Atom One Dark": "atom-one-dark",
        "Monokai Sublime": "monokai-sublime",
        "VS Code Dark": "vs2015",
    }
    highlight_theme = code_theme_map.get(code_theme_name, "github")

    font_map = {
        "Sans-serif": (
            '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica,'
            " Arial, sans-serif"
        ),
        "Serif": 'Georgia, Cambria, "Times New Roman", Times, serif',
        "Monospace": (
            'Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace'
        ),
    }
    chosen_font = font_map.get(font_style, font_style)

    radius_map = {
        "Sharp (0px)": "0px",
        "Rounded (6px)": "6px",
        "Pill (12px)": "12px",
    }
    chosen_radius = radius_map.get(block_radius, "6px")
    chosen_padding = f"{margin_val}mm"
    box_shadow_css = (
        "box-shadow: 0 4px 12px rgba(0,0,0,0.06);" if shadow_enabled else ""
    )
    
    ist_time = datetime.datetime.utcnow() + datetime.timedelta(hours=5, minutes=30)
    timestamp_str = (
        ist_time.strftime("%Y-%m-%d %H:%M IST")
        if show_timestamp
        else ""
    )

    timestamp_html = (
        f'<div style="font-size: 0.55em; font-weight: normal; opacity: 0.7;'
        f' margin-top: 4px;">Exported on: {timestamp_str}</div>'
        if show_timestamp
        else ""
    )

    header_html = (
        f'''
    <div style="font-size: 1.4em; font-weight: bold; text-align: {header_alignment.lower()}; border-bottom: 2px solid {chosen_accent}; padding-bottom: 10px; margin-bottom: 24px; color: {text_color} !important;">
        {custom_title}
        {timestamp_html}
    </div>
    '''
        if (show_header_title and custom_title.strip())
        else ""
    )

    toc_html = ""
    if include_toc and len(active_sections) > 1:
        toc_items = "".join([
            f'<li><a href="#section-{idx}" style="color: {chosen_accent};'
            f' text-decoration: none;">{sec_title}</a></li>'
            for idx, (sec_title, _) in enumerate(active_sections)
        ])
        toc_html = f'''
        <div style="background-color: {code_bg}; border: 1px solid {border_color}; padding: 16px 20px; border-radius: 8px; margin-bottom: 30px;">
            <div style="font-weight: bold; margin-bottom: 8px; font-size: 1.1em;">📑 Table of Contents</div>
            <ul style="margin: 0; padding-left: 20px;">
                {toc_items}
            </ul>
        </div>
        '''

    watermark_html = (
        f'''
    <div style="position: fixed; top: 50%; left: 50%; transform:'''
        " translate(-50%, -50%) rotate(-30deg); font-size: 5rem; font-weight:"
        f" bold; color: {text_color}; opacity: {watermark_alpha};"
        ' pointer-events: none; z-index: 9999; white-space: nowrap;">'
        f" {watermark}"
        " </div>"
    ) if watermark else ""

    page_bottom_right_content = (
        '"Page " counter(page) " of " counter(pages)'
        if include_page_nos
        else '""'
    )

    page_num_css = f'''
    @page {{
        @bottom-right {{
            content: {page_bottom_right_content};
            font-size: 8pt;
            color: #888888;
        }}
    }}
    '''

    footer_wm_html = (
        f'''
        <div style="margin-top: 50px; padding-top: 15px; border-top: 1px solid {border_color}; font-size: 0.85em; text-align: center; color: {text_color}; opacity: 0.8;">
            {footer_wm}
        </div>
        '''
    ) if footer_wm else ""

    divider_html = (
        f'''<hr style="border: none; border-top: 1px solid {border_color};'''
        ' margin: 24px 0;" />'
    ) if show_dividers else ""

    is_chat_bubbles = layout_view_style == "Modern Chat Bubbles"

    document_body_html = ""
    global_msg_idx = 0
    for idx, (sec_title, messages) in enumerate(active_sections):
        section_id = f"section-{idx}"
        if sec_title and sec_title.strip():
            document_body_html += f'''
            <div id="{section_id}" style="font-size: 1.2em; font-weight: bold; color: {chosen_accent}; margin-top: 40px; margin-bottom: 20px; border-bottom: 1px dashed {border_color}; padding-bottom: 8px;">
                {sec_title}
            </div>
            '''
        for msg in messages:
            role = msg["role"]
            
            if global_msg_idx in injected_notes_dict and injected_notes_dict[global_msg_idx].strip():
                note_content = injected_notes_dict[global_msg_idx]
                document_body_html += f'''
                <div style="background: rgba(37, 99, 235, 0.05); border-left: 3px solid #2563eb; padding: 12px 16px; margin: 16px 0; border-radius: 4px; font-style: italic;">
                    <strong>📝 Custom Note:</strong> {note_content}
                </div>
                '''

            if role == "User":
                if is_chat_bubbles:
                    document_body_html += f'''
                    <div style="display: flex; justify-content: flex-end; margin: 16px 0;">
                        <div style="background-color: {user_bg} !important; border: 1px solid {border_color}; padding: 14px 18px; max-width: 85%; border-radius: {chosen_radius} {chosen_radius} 0 {chosen_radius}; color: {user_text} !important; {box_shadow_css}">
                            <div style="font-size: 0.75em; font-weight: bold; margin-bottom: 4px; opacity: 0.8;">You</div>
                            <div style="font-size: {font_size}; color: {user_text} !important;">{msg["html"]}</div>
                        </div>
                    </div>
                    '''
                else:
                    document_body_html += f'''
                    <div style="background-color: {user_bg} !important; border-left: 4px solid {chosen_accent}; padding: 14px 18px; margin: 20px 0; border-radius: 0 {chosen_radius} {chosen_radius} 0; color: {user_text} !important; {box_shadow_css}">
                        <div style="font-size: {font_size}; color: {user_text} !important;">{msg["html"]}</div>
                    </div>
                    {divider_html}
                    '''
            else:
                if is_chat_bubbles:
                    document_body_html += f'''
                    <div style="display: flex; justify-content: flex-start; margin: 16px 0;">
                        <div style="background-color: {code_bg} !important; border: 1px solid {border_color}; padding: 14px 18px; max-width: 85%; border-radius: {chosen_radius} {chosen_radius} {chosen_radius} 0; color: {text_color} !important; {box_shadow_css}">
                            <div style="font-size: 0.75em; font-weight: bold; margin-bottom: 4px; color: {chosen_accent};">ChatGPT</div>
                            <div style="font-size: {font_size}; color: {text_color} !important;">{msg["html"]}</div>
                        </div>
                    </div>
                    '''
                else:
                    document_body_html += f'''
                    <div style="margin: 20px 0; color: {text_color} !important; font-size: {font_size};">
                        {msg["html"]}
                    </div>
                    {divider_html}
                    '''
            global_msg_idx += 1

    if include_study_sheet:
        document_body_html += generate_study_revision_sheet(active_sections)

    footer_html = (
        f'''
    <div style="margin-top: 40px; padding-top: 15px; border-top: 1px dashed'''
        f" {border_color}; font-size: 0.85em; text-align: center; color:"
        f' {text_color}; opacity: 0.8;">'
        f" {footer_text}"
        " </div>"
    ) if footer_text else ""

    full_html = f'''
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/highlight.js/11.9.0/styles/{highlight_theme}.min.css">
        <script src="https://cdnjs.cloudflare.com/ajax/libs/highlight.js/11.9.0/highlight.min.js"></script>
        <script src="https://cdn.jsdelivr.net/npm/mermaid/dist/mermaid.min.js"></script>
        <script>
            hljs.highlightAll();
            mermaid.initialize({{ startOnLoad: true, theme: 'default' }});
        </script>
        <style>
            * {{ box-sizing: border-box; -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }}
            {page_num_css}
            html, body {{
                background-color: {bg_color} !important;
                color: {text_color} !important;
                font-family: {chosen_font};
                font-size: {font_size};
                line-height: {line_height};
                margin: 0;
                padding: 0;
                scroll-behavior: smooth;
            }}
            .page-wrapper {{
                max-width: {max_width}px;
                margin: 0 auto;
                padding: {chosen_padding};
                word-wrap: break-word;
                overflow-wrap: break-word;
            }}
            body *, p, span, h1, h2, h3, h4, h5, h6, li, td, th, strong, b, em, div {{
                color: {text_color} !important;
                font-family: {chosen_font} !important;
            }}
            h1, h2, h3, h4 {{ font-weight: bold; margin-top: 24px; margin-bottom: 12px; }}
            p {{ margin-top: 0; margin-bottom: {paragraph_spacing}px; }}
            ul, ol {{ padding-left: 20px; margin-top: 5px; margin-bottom: 15px; }}
            li {{ margin-bottom: 6px; }}
            pre, pre code {{ background-color: {code_bg} !important; color: {code_color} !important; padding: 14px; border-radius: 6px; white-space: pre-wrap; word-break: break-all; font-family: Consolas, Monaco, monospace; border: 1px solid {border_color}; }}
            code {{ font-family: Consolas, Monaco, monospace; padding: 2px 4px; border-radius: 4px; color: {code_color} !important; background-color: {code_bg} !important; }}
            table {{ border-collapse: collapse; width: 100%; margin-top: 15px; margin-bottom: 20px; table-layout: fixed; }}
            th, td {{ border: 1px solid {border_color}; padding: 10px 14px; text-align: left; word-wrap: break-word; }}
            th {{ background-color: {code_bg} !important; font-weight: bold; color: {text_color} !important; }}
            img {{ max-width: 100%; height: auto; border-radius: 6px; }}
            .mermaid {{ background-color: {code_bg}; padding: 12px; border-radius: 8px; text-align: center; margin: 15px 0; }}
            {custom_css}
        </style>
    </head>
    <body>
        {watermark_html}
        <div class="page-wrapper">
            {header_html}
            {toc_html}
            {document_body_html}
            {footer_html}
            {footer_wm_html if include_footer_wm else ""}
        </div>
    </body>
    </html>
    '''
    return full_html


def convert_chats_to_markdown(
    chat_sections: list,
    custom_title: str,
    message_filter_role: str,
    selection_mode: str,
    allowed_ids: list,
) -> str:
    active_sections = get_filtered_sections(
        chat_sections, message_filter_role, selection_mode, allowed_ids
    )
    md_content = f"# {custom_title}\n\n" if custom_title else ""
    for sec_title, messages in active_sections:
        if sec_title:
            md_content += f"## {sec_title}\n\n"
        for msg in messages:
            md_content += f"### **{msg['role']}**\n\n{msg['html']}\n\n---\n\n"
    return md_content


def convert_chats_to_docx(
    chat_sections: list,
    custom_title: str,
    message_filter_role: str,
    selection_mode: str,
    allowed_ids: list,
    show_timestamp: bool,
    footer_text: str,
    include_page_nos: bool,
    footer_wm: str,
    include_wm: bool,
) -> bytes:
    active_sections = get_filtered_sections(
        chat_sections, message_filter_role, selection_mode, allowed_ids
    )
    doc = Document()

    section = doc.sections[0]
    footer = section.footer
    p = footer.paragraphs[0]

    effective_wm = footer_wm if include_wm else ""
    if effective_wm:
        clean_wm = re.sub(r'<[^>]*?>', '', effective_wm)
        r_wm = p.add_run(clean_wm)
        r_wm.font.size = Pt(8.5)
        r_wm.font.color.rgb = RGBColor(120, 120, 120)
        if include_page_nos:
            p.add_run("\t\t")

    if include_page_nos:
        r_page = p.add_run("Page ")
        r_page.font.size = Pt(8.5)
        r_page.font.color.rgb = RGBColor(120, 120, 120)
        add_docx_page_number(p.add_run())

    if custom_title:
        title_p = doc.add_paragraph()
        title_run = title_p.add_run(custom_title)
        title_run.bold = True
        title_run.font.size = Pt(18)
        title_p.paragraph_format.space_after = Pt(12)

    if show_timestamp:
        ist_time_docx = datetime.datetime.utcnow() + datetime.timedelta(hours=5, minutes=30)
        ts_p = doc.add_paragraph()
        ts_run = ts_p.add_run(
            f"Exported on: {ist_time_docx.strftime('%Y-%m-%d %H:%M IST')}"
        )
        ts_run.font.size = Pt(10)
        ts_run.font.color.rgb = RGBColor(120, 120, 120)
        ts_p.paragraph_format.space_after = Pt(18)

    for sec_title, messages in active_sections:
        if sec_title:
            sec_p = doc.add_paragraph()
            sec_run = sec_p.add_run(sec_title)
            sec_run.bold = True
            sec_run.font.size = Pt(14)
            sec_p.paragraph_format.space_before = Pt(14)
            sec_p.paragraph_format.space_after = Pt(8)

        for msg in messages:
            role = msg["role"]
            role_p = doc.add_paragraph()
            role_run = role_p.add_run(f"{role}:")
            role_run.bold = True
            role_run.font.size = Pt(12)
            role_p.paragraph_format.space_before = Pt(10)
            role_p.paragraph_format.space_after = Pt(4)

            clean_text = re.sub("<[^<]+?>", "", msg["html"])
            msg_p = doc.add_paragraph()
            msg_p.add_run(clean_text)
            msg_p.paragraph_format.space_after = Pt(12)

    if footer_text:
        doc.add_paragraph().paragraph_format.space_after = Pt(24)
        footer_p = doc.add_paragraph()
        f_run = footer_p.add_run(footer_text)
        f_run.font.size = Pt(9)
        f_run.font.color.rgb = RGBColor(100, 100, 100)
        footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        temp_path = tmp.name
        doc.save(temp_path)

    with open(temp_path, "rb") as f:
        docx_bytes = f.read()
    if os.path.exists(temp_path):
        os.remove(temp_path)
    return docx_bytes


# --- MAIN INTERFACE TABS ---
tab_input, tab_custom, tab_preview, tab_download = st.tabs(
    ["🔗 1. Input & Fetch", "📝 2. Custom Notes & Study", "👁️ 3. Live Preview", "📥 4. Download & Audio"]
)

with tab_input:
    st.subheader("Enter ChatGPT Share Links")
    input_mode = st.radio(
        "Select Input Mode:",
        ["Single Link", "Bulk Links (Multiple)"],
        horizontal=True,
    )

    urls = []
    link_titles = {}

    if input_mode == "Single Link":
        share_url_single = st.text_input(
            "ChatGPT Share Link (Supports standard & Custom GPT links)",
            placeholder="https://chatgpt.com/share/... or https://chatgpt.com/g/.../shared/c/..."
        )
        if share_url_single.strip():
            urls = [share_url_single.strip()]

    else:  # Bulk Links (Multiple)
        share_urls_textarea = st.text_area(
            "ChatGPT Share Links (Paste one link per line)",
            placeholder=(
                "https://chatgpt.com/share/...\nhttps://chatgpt.com/g/.../shared/c/..."
            ),
            height=90,
        )
        
        col_load, col_clear = st.columns([2, 5])
        with col_load:
            if st.button("📥 Load Links into Reorder Manager", key="load_bulk"):
                raw_urls = [u.strip() for u in share_urls_textarea.split("\n") if u.strip()]
                st.session_state.bulk_links_data = [
                    {"Position": i + 1, "URL": u, "Title": f"Conversation Section #{i + 1}"}
                    for i, u in enumerate(raw_urls)
                ]

        if st.session_state.bulk_links_data:
            st.markdown("##### ↕️ Bulk Links & Custom Order Manager")
            edited_df = st.data_editor(
                st.session_state.bulk_links_data,
                num_rows="dynamic",
                use_container_width=True,
                key="bulk_order_editor",
                column_config={
                    "Position": st.column_config.NumberColumn("Order #", min_value=1, step=1, format="%d"),
                    "URL": st.column_config.TextColumn("Share URL"),
                    "Title": st.column_config.TextColumn("Section Title")
                }
            )
            
            try:
                sorted_links = sorted(edited_df, key=lambda x: int(x.get("Position", 1)))
            except Exception:
                sorted_links = edited_df

            urls = [row["URL"] for row in sorted_links if row.get("URL", "").strip()]
            for idx, row in enumerate(sorted_links):
                if row.get("URL", "").strip():
                    link_titles[idx] = row.get("Title", f"Conversation Section #{idx + 1}")

    if st.button("🔍 Fetch Chat Content Now", type="primary"):
        if not urls:
            st.error("Please enter and load at least one valid ChatGPT share link.")
        else:
            st.session_state.chat_sections = []
            progress_bar = st.progress(0)
            status_text = st.empty()

            translation_code = lang_code_map.get(target_language, "es") if enable_translation else None

            failed_links = []
            for i, url in enumerate(urls):
                status_text.text(f"Fetching link {i+1} of {len(urls)}...")
                try:
                    msgs = scrape_single_chat(url, enable_ai_filter, translation_code)
                    if msgs:
                        sec_title = (
                            link_titles.get(i, f"Conversation Section #{i+1}")
                            if input_mode == "Bulk Links (Multiple)"
                            else ""
                        )
                        st.session_state.chat_sections.append((sec_title, msgs))
                except Exception as e:
                    failed_links.append(f"{url} (Error: {e})")
                progress_bar.progress((i + 1) / len(urls))

            status_text.empty()
            progress_bar.empty()

            if st.session_state.chat_sections:
                st.success(
                    f"Successfully fetched {len(st.session_state.chat_sections)}"
                    " conversation(s)! Head over to the **Live Preview** tab."
                )
                if failed_links:
                    st.warning(
                        f"Failed to load {len(failed_links)} link(s):\n"
                        + "\n".join(failed_links)
                    )
            else:
                st.error(
                    "Could not parse any conversations from the provided links. Please check the URLs."
                )

all_message_ids = []
id_to_label_map = {}
flat_messages_list = []
total_word_count = 0
if st.session_state.chat_sections:
    for _, msgs in st.session_state.chat_sections:
        for m in msgs:
            mid = m["id"]
            all_message_ids.append(mid)
            id_to_label_map[mid] = f"[{m['role']}] {m['text'][:55]}..."
            flat_messages_list.append(m)
            total_word_count += len(m["text"].split())

    st.sidebar.markdown("---")
    st.sidebar.markdown("📊 **Document Stats**")
    st.sidebar.caption(
        f"• Total Words: **{total_word_count:,}**\n• Est. Reading Time:"
        f" **{max(1, total_word_count // 200)} min**"
    )

selected_message_ids = all_message_ids
if message_selection_mode == "Pick Specific Messages (Custom)" and all_message_ids:
    st.sidebar.markdown("---")
    st.sidebar.markdown("##### 🎯 Custom Message Selector")
    selected_message_ids = st.sidebar.multiselect(
        "Check messages to include:",
        options=all_message_ids,
        default=all_message_ids,
        format_func=lambda x: id_to_label_map.get(x, x),
    )

with tab_custom:
    st.subheader("📝 Custom Note Injector Manager")
    if enable_custom_notes:
        st.markdown("Insert custom personal notes between specific imported conversation turns.")
        if st.session_state.chat_sections and flat_messages_list:
            selected_idx_for_note = st.selectbox(
                "Select position (insert note right before message):",
                options=list(range(len(flat_messages_list))),
                format_func=lambda i: f"Before [{flat_messages_list[i]['role']}] {flat_messages_list[i]['text'][:45]}..."
            )
            
            note_text_input = st.text_area("Custom Note Content (Markdown supported)", placeholder="Type your personal note or comment here...")
            
            if st.button("➕ Inject Note at Position"):
                st.session_state.injected_notes[selected_idx_for_note] = note_text_input
                st.success("Note successfully added! It will appear in the live preview.")
                
            if st.session_state.injected_notes:
                st.markdown("##### Current Injected Notes:")
                for idx, ntext in list(st.session_state.injected_notes.items()):
                    col_n1, col_n2 = st.columns([5, 1])
                    with col_n1:
                        st.write(f"• **Pos {idx}**: {ntext}")
                    with col_n2:
                        if st.button("Remove", key=f"del_note_{idx}"):
                            del st.session_state.injected_notes[idx]
                            st.rerun()
        else:
            st.info("Fetch chat content first in Tab 1 to enable note injection.")
    else:
        st.info("The **Custom Note Injector** toggle is currently turned OFF in the sidebar. Switch it on to use this feature.")

with tab_preview:
    st.subheader("Live Document Preview")
    if st.session_state.chat_sections:
        st.caption(
            "Instantly reflects your chosen theme, custom titles, images, Mermaid diagrams, and revision sheets."
        )
        live_preview_html = build_html_document(
            st.session_state.chat_sections,
            theme,
            layout_view_style,
            font_style,
            font_size,
            line_height,
            paragraph_spacing,
            message_filter_role,
            message_selection_mode,
            selected_message_ids,
            show_header_title,
            custom_title,
            header_alignment,
            show_dividers,
            show_timestamp,
            footer_text,
            chosen_accent,
            code_theme,
            block_radius,
            margin_mm,
            max_content_width,
            enable_box_shadow,
            generate_toc,
            watermark_text,
            watermark_opacity / 100.0,
            footer_watermark_text,
            show_page_numbers,
            custom_css_injection,
            enable_study_revision,
            st.session_state.injected_notes if enable_custom_notes else {},
        )
        with st.container(border=True):
            st.components.v1.html(live_preview_html, height=700, scrolling=True)
    else:
        st.info(
            "Please enter your share links and click **Fetch Chat Content Now** in the first tab to view the preview."
        )

with tab_download:
    st.subheader("Download Document File & Audio Export")
    if st.session_state.chat_sections:
        output_filename = st.text_input("Output Filename", value="chatgpt_export")

        col_d1, col_d2 = st.columns(2)
        with col_d1:
            if st.button("🚀 Generate & Download File Now", type="primary"):
                try:
                    base_filename = output_filename.strip()
                    if not base_filename:
                        base_filename = "chat_export"

                    current_html = build_html_document(
                        st.session_state.chat_sections,
                        theme,
                        layout_view_style,
                        font_style,
                        font_size,
                        line_height,
                        paragraph_spacing,
                        message_filter_role,
                        message_selection_mode,
                        selected_message_ids,
                        show_header_title,
                        custom_title,
                        header_alignment,
                        show_dividers,
                        show_timestamp,
                        footer_text,
                        chosen_accent,
                        code_theme,
                        block_radius,
                        margin_mm,
                        max_content_width,
                        enable_box_shadow,
                        generate_toc,
                        watermark_text,
                        watermark_opacity / 100.0,
                        footer_watermark_text,
                        show_page_numbers,
                        custom_css_injection,
                        enable_study_revision,
                        st.session_state.injected_notes if enable_custom_notes else {},
                    )

                    if "PDF" in export_format:
                        with st.spinner("Compiling high-fidelity PDF Document..."):
                            file_data = build_pdf_document(current_html, margin_val=margin_mm)
                            mime_type = "application/pdf"
                            final_name = f"{base_filename}.pdf"
                    elif "Word" in export_format:
                        with st.spinner("Compiling Word Document (.docx)..."):
                            file_data = convert_chats_to_docx(
                                st.session_state.chat_sections,
                                custom_title,
                                message_filter_role,
                                message_selection_mode,
                                selected_message_ids,
                                show_timestamp,
                                footer_text,
                                show_page_numbers,
                                footer_watermark_text,
                                include_footer_wm,
                            )
                            mime_type = (
                                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )
                            final_name = f"{base_filename}.docx"
                    elif "Markdown" in export_format:
                        with st.spinner("Generating Markdown file..."):
                            md_text = convert_chats_to_markdown(
                                st.session_state.chat_sections,
                                custom_title,
                                message_filter_role,
                                message_selection_mode,
                                selected_message_ids,
                            )
                            file_data = md_text.encode("utf-8")
                            mime_type = "text/markdown"
                            final_name = f"{base_filename}.md"
                    else:  # HTML
                        with st.spinner("Generating HTML file..."):
                            file_data = current_html.encode("utf-8")
                            mime_type = "text/html"
                            final_name = f"{base_filename}.html"

                    st.success("Document built successfully! Click below to save:")
                    st.download_button(
                        label=f"💾 Click Here to Save {final_name}",
                        data=file_data,
                        file_name=final_name,
                        mime=mime_type,
                    )
                except Exception as e:
                    st.error(f"An error occurred during file creation: {e}")

        with col_d2:
            if HAS_GTTS:
                st.markdown("##### 🎧 Text-to-Speech Audio Summary (.mp3)")
                if st.button("🔊 Generate Audio Podcast Summary"):
                    with st.spinner("Synthesizing audio file using TTS..."):
                        try:
                            summary_corpus = f"Audio summary for document: {custom_title}. "
                            for _, msgs in st.session_state.chat_sections:
                                for m in msgs:
                                    if m["role"] == "ChatGPT":
                                        summary_corpus += m["text"][:300] + ". "

                            tts = gTTS(text=summary_corpus[:3500], lang='en', slow=False)
                            audio_tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".mp3")
                            tts.save(audio_tmp.name)
                            
                            with open(audio_tmp.name, "rb") as af:
                                audio_bytes = af.read()
                            os.remove(audio_tmp.name)

                            st.audio(audio_bytes, format="audio/mp3")
                            st.download_button(
                                label="📥 Download Audio (.mp3)",
                                data=audio_bytes,
                                file_name=f"{output_filename}_summary.mp3",
                                mime="audio/mp3"
                            )
                        except Exception as tts_err:
                            st.error(f"TTS generation failed: {tts_err}")
            else:
                st.info("Install `gTTS` library to enable MP3 audio exports.")
    else:
        st.info("Fetch your chats first before downloading.")