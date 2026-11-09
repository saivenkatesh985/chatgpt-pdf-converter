import os
import tempfile
import datetime
import re
import requests
import markdown
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Page config
st.set_page_config(page_title="ChatGPT Document Customizer", page_icon="📄", layout="wide")

# Custom UI polish
st.markdown("""
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
    </style>
""", unsafe_allow_html=True)

st.title("📄 ChatGPT Link to Multi-Format Document & Preview")
st.markdown("Transform single or multiple public ChatGPT conversations into professionally styled Word, Markdown, or HTML documents with advanced customization options!")

# --- Sidebar Customization Controls ---
st.sidebar.header("🎨 Customization & Layout Studio")

theme = st.sidebar.selectbox("Color Theme", ["White", "Dark", "Note", "Study"])
font_style = st.sidebar.selectbox("Font Style", ["Sans-serif", "Serif", "Monospace"])

font_size = st.sidebar.text_input("Font Size (e.g., 14px, 16px, 1.1rem)", value="16px")
margin_mm = st.sidebar.number_input("Page Margins (mm)", min_value=0, max_value=50, value=15, step=1)
max_content_width = st.sidebar.slider("Max Content Width (px)", min_value=600, max_value=1400, value=900, step=50)

st.sidebar.markdown("---")
st.sidebar.header("⚙️ Advanced Styling & Structure")

export_format = st.sidebar.selectbox("Export Format Type", ["Word Document (.docx)", "Markdown Document (.md)", "HTML Webpage (.html)"])
show_user_messages = st.sidebar.checkbox("Include User Messages", value=True)
remove_ai_closings = st.sidebar.checkbox("Remove AI Closing Offers ('If you want, I can...')", value=True)
show_header_title = st.sidebar.checkbox("Include Document Title Header", value=True)
custom_title = st.sidebar.text_input("Document Title Text", value="ChatGPT Conversation Export")
header_alignment = st.sidebar.selectbox("Header Alignment", ["Left", "Center", "Right"])
show_dividers = st.sidebar.checkbox("Show Turn Dividers", value=False)
show_timestamp = st.sidebar.checkbox("Include Export Timestamp", value=True)
footer_text = st.sidebar.text_input("Custom Footer Text", value="")

accent_color_map = {
    "Modern Green (#10a37f)": "#10a37f",
    "Executive Blue (#2563eb)": "#2563eb",
    "Royal Purple (#7c3aed)": "#7c3aed",
    "Sunset Orange (#ea580c)": "#ea580c"
}
chosen_accent_label = st.sidebar.selectbox("Accent & Border Color", list(accent_color_map.keys()))
chosen_accent = accent_color_map[chosen_accent_label]

line_height = st.sidebar.slider("Line Spacing", min_value=1.2, max_value=2.0, value=1.6, step=0.1)
paragraph_spacing = st.sidebar.slider("Paragraph Spacing (px)", min_value=4, max_value=24, value=12, step=2)
block_radius = st.sidebar.selectbox("User Message Box Corners", ["Sharp (0px)", "Rounded (6px)", "Pill (12px)"])
enable_box_shadow = st.sidebar.checkbox("Add Subtle Shadow to User Boxes", value=True)

# --- Main Form Mode Selector & Inputs ---
input_mode = st.radio("Select Input Mode:", ["Single Link", "Bulk Links (Multiple)"], horizontal=True)

bulk_order = "Chronological (Top-to-Bottom)"
link_titles = {}
urls = []

if input_mode == "Single Link":
    share_url_single = st.text_input("ChatGPT Share Link", placeholder="https://chatgpt.com/share/...")
    if share_url_single.strip():
        urls = [share_url_single.strip()]
else:
    share_urls_textarea = st.text_area("ChatGPT Share Links (Paste one link per line)", placeholder="https://chatgpt.com/share/...\nhttps://chatgpt.com/share/...", height=100)
    bulk_order = st.radio("Bulk Chat Sequence Order:", ["Chronological (Top-to-Bottom)", "Reverse Chronological (Bottom-to-Top)"], horizontal=True)
    
    raw_urls = [u.strip() for u in share_urls_textarea.split("\n") if u.strip()]
    if raw_urls:
        if bulk_order.startswith("Reverse"):
            raw_urls = raw_urls[::-1]
            
        st.markdown("##### ✏️ Custom Section Titles per Link")
        for idx, u in enumerate(raw_urls):
            link_titles[u] = st.text_input(f"Title for Link #{idx+1} ({u[:30]}...)", value=f"Conversation Section #{idx+1}")
    urls = raw_urls

output_filename = st.text_input("Output Filename", value="chatgpt_export")

def clean_assistant_text(text: str) -> str:
    """Removes trailing AI closing offers and follow-up bullets."""
    pattern = re.compile(r'\n+(?:If you want, I can:|Would you like me to:|Let me know if you).*$', re.DOTALL | re.IGNORECASE)
    return pattern.sub('', text).strip()

def scrape_single_chat(url: str, strip_closings: bool) -> list:
    """Fetches chat content via ChatGPT's backend share API and parses Markdown into formatted HTML."""
    clean_url = url.split("?")[0].strip()
    if not clean_url:
        return []
    
    if "/g/" in clean_url or "g-p-" in clean_url:
        raise Exception(f"Custom GPT links require login: {clean_url}")

    parts = clean_url.rstrip("/").split("/")
    share_id = parts[-1]
    
    if not share_id or len(share_id) < 10:
        raise Exception(f"Invalid share link format: {clean_url}")

    api_endpoint = f"https://chatgpt.com/backend-api/share/{share_id}"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
        "Accept": "application/json"
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
                    content_parts = message.get("content", {}).get("parts", [])
                    text_content = "".join([str(p) for p in content_parts if isinstance(p, str)])
                    if text_content.strip():
                        role_label = "User" if author == "user" else "ChatGPT"
                        
                        if role_label == "ChatGPT" and strip_closings:
                            text_content = clean_assistant_text(text_content)
                            
                        if text_content.strip():
                            formatted_html = markdown.markdown(text_content, extensions=['fenced_code', 'tables'])
                            messages.append({
                                "role": role_label,
                                "html": formatted_html
                            })
    except Exception as parse_err:
        raise Exception(f"Error parsing conversation JSON: {parse_err}")

    if not messages:
        raise Exception("No messages found in this share link.")

    return messages

def build_html_document(chat_sections: list, theme: str, font_style: str, font_size: str, line_height: float, paragraph_spacing: int, show_user_messages: bool, show_header_title: bool, custom_title: str, header_alignment: str, show_dividers: bool, show_timestamp: bool, footer_text: str, chosen_accent: str, block_radius: str, margin_val: int, max_width: int, shadow_enabled: bool) -> str:
    """Builds a combined HTML document using custom numeric margin padding and layout settings."""
    if theme == "Dark":
        bg_color, text_color = "#121212", "#e0e0e0"
        user_bg, user_text = "#1a332c", "#a3e4d7"
        border_color, code_bg, code_color = "#333333", "#0f0f0f", "#f8f8f8"
    elif theme == "Note":
        bg_color, text_color = "#fbf0d9", "#2c221e"
        user_bg, user_text = "#e6f2ed", "#1b4d3e"
        border_color, code_bg, code_color = "#e6d7bc", "#f4e7ce", "#2c221e"
    elif theme == "Study":
        bg_color, text_color = "#f0f4f8", "#102a43"
        user_bg, user_text = "#d9e2ec", "#243b53"
        border_color, code_bg, code_color = "#bcccdc", "#e2e8f0", "#102a43"
    else:
        bg_color, text_color = "#ffffff", "#1a1a1a"
        user_bg, user_text = "#edf7f0", "#114f2e"
        border_color, code_bg, code_color = "#e1e4e8", "#f1f1f1", "#222222"

    font_map = {
        "Sans-serif": '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif',
        "Serif": 'Georgia, Cambria, "Times New Roman", Times, serif',
        "Monospace": 'Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace'
    }
    chosen_font = font_map.get(font_style, font_map["Sans-serif"])

    radius_map = {
        "Sharp (0px)": "0px",
        "Rounded (6px)": "6px",
        "Pill (12px)": "12px"
    }
    chosen_radius = radius_map.get(block_radius, "6px")
    chosen_padding = f"{margin_val}mm"
    box_shadow_css = "box-shadow: 0 4px 12px rgba(0,0,0,0.06);" if shadow_enabled else ""

    timestamp_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M") if show_timestamp else ""

    header_html = f"""
    <div style="font-size: 1.4em; font-weight: bold; text-align: {header_alignment.lower()}; border-bottom: 2px solid {chosen_accent}; padding-bottom: 10px; margin-bottom: 24px; color: {text_color} !important;">
        {custom_title}
        {f'<div style="font-size: 0.55em; font-weight: normal; opacity: 0.7; margin-top: 4px;">Exported on: {timestamp_str}</div>' if show_timestamp else ''}
    </div>
    """ if show_header_title else ""

    divider_html = f"""<hr style="border: none; border-top: 1px solid {border_color}; margin: 24px 0;" />""" if show_dividers else ""

    document_body_html = ""
    for sec_title, messages in chat_sections:
        if len(chat_sections) > 1 or sec_title:
            document_body_html += f"""
            <div style="font-size: 1.2em; font-weight: bold; color: {chosen_accent}; margin-top: 40px; margin-bottom: 20px; border-bottom: 1px dashed {border_color}; padding-bottom: 8px;">
                {sec_title}
            </div>
            """
        for msg in messages:
            role = msg["role"]
            if role == "User":
                if not show_user_messages:
                    continue
                document_body_html += f"""
                <div style="background-color: {user_bg} !important; border-left: 4px solid {chosen_accent}; padding: 14px 18px; margin: 20px 0; border-radius: 0 {chosen_radius} {chosen_radius} 0; color: {user_text} !important; {box_shadow_css}">
                    <div style="font-size: {font_size}; color: {user_text} !important;">{msg["html"]}</div>
                </div>
                {divider_html}
                """
            else:
                document_body_html += f"""
                <div style="margin: 20px 0; color: {text_color} !important; font-size: {font_size};">
                    {msg["html"]}
                </div>
                {divider_html}
                """

    footer_html = f"""
    <div style="margin-top: 40px; padding-top: 15px; border-top: 1px dashed {border_color}; font-size: 0.85em; text-align: center; color: {text_color}; opacity: 0.8;">
        {footer_text}
    </div>
    """ if footer_text else ""

    full_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            * {{ box-sizing: border-box; -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }}
            html, body {{
                background-color: {bg_color} !important;
                color: {text_color} !important;
                font-family: {chosen_font};
                font-size: {font_size};
                line-height: {line_height};
                margin: 0;
                padding: 0;
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
            pre {{ background-color: {code_bg} !important; color: {code_color} !important; padding: 12px; border-radius: 6px; white-space: pre-wrap; word-break: break-all; font-family: Consolas, Monaco, monospace; }}
            code {{ font-family: Consolas, Monaco, monospace; background-color: {code_bg} !important; color: {code_color} !important; padding: 2px 4px; border-radius: 4px; }}
            table {{ border-collapse: collapse; width: 100%; margin-top: 15px; margin-bottom: 20px; table-layout: fixed; }}
            th, td {{ border: 1px solid {border_color}; padding: 10px 14px; text-align: left; word-wrap: break-word; }}
            th {{ background-color: {code_bg} !important; font-weight: bold; }}
            img {{ max-width: 100%; height: auto; border-radius: 6px; }}
        </style>
    </head>
    <body>
        <div class="page-wrapper">
            {header_html}
            {document_body_html}
            {footer_html}
        </div>
    </body>
    </html>
    """
    return full_html

def convert_chats_to_markdown(chat_sections: list, custom_title: str) -> str:
    md_content = f"# {custom_title}\n\n"
    for sec_title, messages in chat_sections:
        if sec_title:
            md_content += f"## {sec_title}\n\n"
        for msg in messages:
            md_content += f"### **{msg['role']}**\n\n{msg['html']}\n\n---\n\n"
    return md_content

def convert_chats_to_docx(chat_sections: list, custom_title: str, show_user_messages: bool, show_timestamp: bool, footer_text: str) -> bytes:
    doc = Document()
    if custom_title:
        title_p = doc.add_paragraph()
        title_run = title_p.add_run(custom_title)
        title_run.bold = True
        title_run.font.size = Pt(18)
        title_p.paragraph_format.space_after = Pt(12)
        
    if show_timestamp:
        ts_p = doc.add_paragraph()
        ts_run = ts_p.add_run(f"Exported on: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}")
        ts_run.font.size = Pt(10)
        ts_run.font.color.rgb = RGBColor(120, 120, 120)
        ts_p.paragraph_format.space_after = Pt(18)

    for sec_title, messages in chat_sections:
        if sec_title:
            sec_p = doc.add_paragraph()
            sec_run = sec_p.add_run(sec_title)
            sec_run.bold = True
            sec_run.font.size = Pt(14)
            sec_p.paragraph_format.space_before = Pt(14)
            sec_p.paragraph_format.space_after = Pt(8)

        for msg in messages:
            role = msg["role"]
            if role == "User" and not show_user_messages:
                continue
            role_p = doc.add_paragraph()
            role_run = role_p.add_run(f"{role}:")
            role_run.bold = True
            role_run.font.size = Pt(12)
            role_p.paragraph_format.space_before = Pt(10)
            role_p.paragraph_format.space_after = Pt(4)
            
            clean_text = re.sub('<[^<]+?>', '', msg['html'])
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

# --- Session State Management ---
if "chat_sections" not in st.session_state:
    st.session_state.chat_sections = None

# Step 1: Action Button to Fetch Chats (Single or Bulk)
if st.button("🔍 1. Fetch Chat Content", type="primary"):
    if not urls:
        st.error("Please enter at least one valid ChatGPT share link.")
    else:
        st.session_state.chat_sections = []
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        failed_links = []
        for i, url in enumerate(urls):
            status_text.text(f"Fetching link {i+1} of {len(urls)}...")
            try:
                msgs = scrape_single_chat(url, remove_ai_closings)
                if msgs:
                    sec_title = link_titles.get(url, f"Conversation Section #{i+1}") if input_mode == "Bulk Links (Multiple)" else (custom_title if show_header_title else "Conversation")
                    st.session_state.chat_sections.append((sec_title, msgs))
            except Exception as e:
                failed_links.append(f"{url} (Error: {e})")
            progress_bar.progress((i + 1) / len(urls))
            
        status_text.empty()
        progress_bar.empty()
        
        if st.session_state.chat_sections:
            st.success(f"Successfully fetched {len(st.session_state.chat_sections)} conversation(s)!")
            if failed_links:
                st.warning(f"Failed to load {len(failed_links)} link(s):\n" + "\n".join(failed_links))
        else:
            st.error("Could not parse any conversations from the provided links. Please check the URLs.")

# Step 2: Stacked Preview and Download Sections
if st.session_state.chat_sections:
    st.markdown("---")
    
    live_preview_html = build_html_document(
        st.session_state.chat_sections, theme, font_style, font_size, line_height, paragraph_spacing,
        show_user_messages, show_header_title, custom_title, header_alignment, 
        show_dividers, show_timestamp, footer_text, chosen_accent, block_radius, margin_mm, max_content_width, enable_box_shadow
    )
    
    # --- LIVE PREVIEW CONTAINER ---
    st.subheader("👁️ 2. Document Preview")
    st.caption("Instantly reflects your chosen theme, custom link titles, and layout configurations.")
    
    with st.container(border=True):
        st.components.v1.html(live_preview_html, height=650, scrolling=True)

    st.markdown("---")
    
    # --- DOWNLOAD SECTION ---
    st.subheader("📥 3. Download Document File")
    st.caption("Save your conversation export in your preferred format.")
    
    if st.button("🚀 Generate & Download File Now", type="secondary"):
        try:
            base_filename = output_filename.strip()
            if not base_filename:
                base_filename = "chatgpt_export"
            
            if "Word" in export_format:
                with st.spinner("Compiling Word Document (.docx)..."):
                    file_data = convert_chats_to_docx(st.session_state.chat_sections, custom_title, show_user_messages, show_timestamp, footer_text)
                    mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    final_name = f"{base_filename}.docx"
            elif "Markdown" in export_format:
                with st.spinner("Generating Markdown file..."):
                    md_text = convert_chats_to_markdown(st.session_state.chat_sections, custom_title)
                    file_data = md_text.encode("utf-8")
                    mime_type = "text/markdown"
                    final_name = f"{base_filename}.md"
            else: # HTML
                with st.spinner("Generating HTML file..."):
                    file_data = live_preview_html.encode("utf-8")
                    mime_type = "text/html"
                    final_name = f"{base_filename}.html"
            
            st.success("Document built successfully! Click below to save:")
            st.download_button(
                label=f"💾 Click Here to Save {final_name}",
                data=file_data,
                file_name=final_name,
                mime=mime_type
            )
        except Exception as e:
            st.error(f"An error occurred during file creation: {e}")