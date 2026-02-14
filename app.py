"""
Print Automation Program
A Streamlit web application for mail merge - generating personalized Word documents
from templates and data files.
"""

import streamlit as st
import zipfile
from io import BytesIO
from utils.document_processor import DocumentProcessor
from utils.data_handler import DataHandler
from utils.email_handler import EmailHandler
from utils.zoho_sign_handler import ZohoSignHandler
import time
import base64
import smtplib
from email.message import EmailMessage

# Page configuration
st.set_page_config(
    page_title="AutoDispatch - Document Generator",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom CSS for premium look
st.markdown(
    """
<style>
    /* Main container styling */
    .main {
        background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
    }
    
    /* Header styling */
    .main-header {
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-size: 2.5rem;
        font-weight: 700;
        text-align: center;
        margin-bottom: 0.5rem;
    }
    
    .sub-header {
        color: #a0a0a0;
        text-align: center;
        font-size: 1.1rem;
        margin-bottom: 2rem;
    }
    
    /* Card styling */
    .upload-card {
        background: linear-gradient(145deg, #1e1e30 0%, #2a2a40 100%);
        border-radius: 16px;
        padding: 1.5rem;
        border: 1px solid rgba(255,255,255,0.1);
        box-shadow: 0 8px 32px rgba(0,0,0,0.3);
        margin-bottom: 1rem;
    }
    
    .card-title {
        color: #ffffff;
        font-size: 1.2rem;
        font-weight: 600;
        margin-bottom: 1rem;
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }
    
    /* Status badges */
    .status-badge {
        display: inline-block;
        padding: 0.25rem 0.75rem;
        border-radius: 20px;
        font-size: 0.85rem;
        font-weight: 500;
    }
    
    .status-success {
        background: rgba(46, 213, 115, 0.2);
        color: #2ed573;
        border: 1px solid rgba(46, 213, 115, 0.3);
    }
    
    .status-warning {
        background: rgba(255, 193, 7, 0.2);
        color: #ffc107;
        border: 1px solid rgba(255, 193, 7, 0.3);
    }
    
    .status-info {
        background: rgba(102, 126, 234, 0.2);
        color: #667eea;
        border: 1px solid rgba(102, 126, 234, 0.3);
    }
    
    /* Placeholder tags */
    .placeholder-tag {
        display: inline-block;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 0.3rem 0.8rem;
        border-radius: 20px;
        margin: 0.2rem;
        font-size: 0.85rem;
        font-weight: 500;
    }
    
    /* Info box */
    .info-box {
        background: rgba(102, 126, 234, 0.1);
        border-left: 4px solid #667eea;
        padding: 1rem;
        border-radius: 0 8px 8px 0;
        margin: 1rem 0;
    }
    
    /* Success box */
    .success-box {
        background: rgba(46, 213, 115, 0.1);
        border-left: 4px solid #2ed573;
        padding: 1rem;
        border-radius: 0 8px 8px 0;
        margin: 1rem 0;
    }
    
    /* Button styling */
    .stButton > button {
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        padding: 0.75rem 2rem;
        border-radius: 8px;
        font-weight: 600;
        transition: all 0.3s ease;
    }
    
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 20px rgba(102, 126, 234, 0.4);
    }
    
    /* Download button */
    .stDownloadButton > button {
        background: linear-gradient(90deg, #2ed573 0%, #1abc9c 100%);
        color: white;
        border: none;
        padding: 0.75rem 2rem;
        border-radius: 8px;
        font-weight: 600;
        width: 100%;
    }
    
    /* Selectbox styling */
    .stSelectbox > div > div {
        background: rgba(255,255,255,0.05);
        border: 1px solid rgba(255,255,255,0.1);
    }
    
    /* Progress styling */
    .stProgress > div > div > div {
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
    }
    
    /* Expander styling */
    .streamlit-expanderHeader {
        background: rgba(255,255,255,0.05);
        border-radius: 8px;
    }
    
    /* Navigation buttons container */
    .nav-buttons {
        display: flex;
        justify-content: space-between;
        margin-top: 2rem;
        padding-top: 1.5rem;
        border-top: 1px solid rgba(255,255,255,0.1);
    }
    
    /* Back button styling */
    div[data-testid="column"]:first-child .stButton > button {
        background: linear-gradient(90deg, #6c757d 0%, #495057 100%);
    }
    
    /* Next button styling - green gradient */
    div[data-testid="column"]:last-child .stButton > button {
        background: linear-gradient(90deg, #2ed573 0%, #1abc9c 100%);
    }
    
    /* Step indicator */
    .step-indicator {
        display: flex;
        justify-content: center;
        gap: 2rem;
        margin-bottom: 2rem;
    }
    
    .step {
        display: flex;
        align-items: center;
        gap: 0.5rem;
        color: #666;
    }
    
    .step.active {
        color: #667eea;
    }
    
    .step.completed {
        color: #2ed573;
    }
    
    .step-number {
        width: 30px;
        height: 30px;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        font-weight: 600;
        border: 2px solid currentColor;
    }
    
    /* Hide Streamlit branding */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    
    /* Dataframe styling */
    .dataframe {
        font-size: 0.85rem;
    }
</style>
""",
    unsafe_allow_html=True,
)


def init_session_state():
    """Initialize session state variables."""
    if "template_processor" not in st.session_state:
        st.session_state.template_processor = None
    if "data_handler" not in st.session_state:
        st.session_state.data_handler = None
    if "column_mapping" not in st.session_state:
        st.session_state.column_mapping = {}
    if "generated_docs" not in st.session_state:
        st.session_state.generated_docs = None
    if "step" not in st.session_state:
        st.session_state.step = 1
    if "current_tab" not in st.session_state:
        st.session_state.current_tab = 0
    
    # Email-related session state
    if "email_handler" not in st.session_state:
        st.session_state.email_handler = None
    if "email_column" not in st.session_state:
        st.session_state.email_column = None
    if "email_validation_results" not in st.session_state:
        st.session_state.email_validation_results = None
    if "missing_emails" not in st.session_state:
        st.session_state.missing_emails = {}
    if "skip_rows" not in st.session_state:
        st.session_state.skip_rows = set()
    if "email_subject_template" not in st.session_state:
        st.session_state.email_subject_template = ""
    if "email_body_template" not in st.session_state:
        st.session_state.email_body_template = ""
    if "smtp_configured" not in st.session_state:
        st.session_state.smtp_configured = False
    if "cc_emails" not in st.session_state:
        st.session_state.cc_emails = ""
    if "bcc_emails" not in st.session_state:
        st.session_state.bcc_emails = ""
    if "common_attachment" not in st.session_state:
        st.session_state.common_attachment = None
    if "email_send_results" not in st.session_state:
        st.session_state.email_send_results = None


def go_to_tab(tab_index):
    """Navigate to a specific tab."""
    st.session_state.current_tab = tab_index


def render_nav_buttons(
    current_tab,
    can_proceed=True,
    show_back=True,
    show_next=True,
    next_label="Next Step ➡️",
):
    """Render navigation buttons at the bottom of each section."""
    st.markdown("---")

    col1, col2, col3 = st.columns([1, 2, 1])

    with col1:
        if show_back and current_tab > 0:
            if st.button("⬅️ Back", key=f"back_{current_tab}", use_container_width=True):
                go_to_tab(current_tab - 1)
                st.rerun()

    with col3:
        if show_next and current_tab < 4:
            if can_proceed:
                if st.button(
                    next_label, key=f"next_{current_tab}", use_container_width=True
                ):
                    go_to_tab(current_tab + 1)
                    st.rerun()
            else:
                st.button(
                    next_label,
                    key=f"next_{current_tab}",
                    use_container_width=True,
                    disabled=True,
                )


def render_header():
    """Render the application header."""
    st.markdown(
        '<h1 class="main-header">📄 AutoDispatch</h1>', unsafe_allow_html=True
    )
    st.markdown(
        '<p class="sub-header">Generate personalized documents from templates and data files</p>',
        unsafe_allow_html=True,
    )


    # Step indicator
    step = st.session_state.step
    st.markdown(
        f"""
    <div class="step-indicator">
        <div class="step {"completed" if step > 1 else "active" if step == 1 else ""}">
            <div class="step-number">{"✓" if step > 1 else "1"}</div>
            <span>Upload Template</span>
        </div>
        <div class="step {"completed" if step > 2 else "active" if step == 2 else ""}">
            <div class="step-number">{"✓" if step > 2 else "2"}</div>
            <span>Upload Data</span>
        </div>
        <div class="step {"completed" if step > 3 else "active" if step == 3 else ""}">
            <div class="step-number">{"✓" if step > 3 else "3"}</div>
            <span>Map Columns</span>
        </div>
        <div class="step {"completed" if step > 4 else "active" if step == 4 else ""}">
            <div class="step-number">{"✓" if step > 4 else "4"}</div>
            <span>Generate</span>
        </div>
        <div class="step {"completed" if step > 5 else "active" if step == 5 else ""}">
            <div class="step-number">{"✓" if step > 5 else "5"}</div>
            <span>Send Emails</span>
        </div>
    </div>
    """,
        unsafe_allow_html=True,
    )


def render_template_upload():
    """Render template upload section."""
    st.markdown("### 📝 Step 1: Upload Word Template")

    st.markdown(
        """
    <div class="info-box">
        <strong>Template Format:</strong> Upload a Word document (.docx) with placeholders in <code>{column_name}</code> format.
        <br><br>
        <strong>Example:</strong> "Dear <code>{Customer Name}</code>, your order <code>{Order ID}</code> is ready."
    </div>
    """,
        unsafe_allow_html=True,
    )

    template_file = st.file_uploader(
        "Upload Word Template",
        type=["docx"],
        help="Upload a .docx file with {placeholder} markers",
    )

    can_proceed = False

    if template_file:
        try:
            template_bytes = template_file.read()
            processor = DocumentProcessor(template_bytes)
            st.session_state.template_processor = processor

            placeholders = processor.get_placeholders()

            if placeholders:
                st.markdown(
                    """
                <div class="success-box">
                    <strong>✅ Template loaded successfully!</strong>
                </div>
                """,
                    unsafe_allow_html=True,
                )

                st.markdown("**Detected Placeholders:**")
                placeholder_html = " ".join(
                    [
                        f'<span class="placeholder-tag">{{{p}}}</span>'
                        for p in placeholders
                    ]
                )
                st.markdown(
                    f'<div style="margin: 0.5rem 0;">{placeholder_html}</div>',
                    unsafe_allow_html=True,
                )

                st.session_state.step = max(st.session_state.step, 2)
                can_proceed = True
            else:
                st.warning(
                    "⚠️ No placeholders found in the template. Make sure to use {column_name} format."
                )

        except Exception as e:
            st.error(f"❌ Error loading template: {str(e)}")
            st.session_state.template_processor = None

    # Navigation buttons
    render_nav_buttons(
        0, can_proceed=can_proceed, show_back=False, next_label="Next: Upload Data ➡️"
    )


def render_data_upload():
    """Render data file upload section."""
    if st.session_state.template_processor is None:
        st.info(
            "👆 Please upload a template first. Click 'Back' to go to the previous step."
        )
        render_nav_buttons(1, can_proceed=False, show_next=False)
        return

    st.markdown("### 📊 Step 2: Upload Data File")

    st.markdown(
        """
    <div class="info-box">
        <strong>Supported Formats:</strong> CSV (.csv), Excel (.xlsx, .xls)
        <br><br>
        <strong>Tip:</strong> Make sure your column names match the placeholders in your template.
    </div>
    """,
        unsafe_allow_html=True,
    )

    data_file = st.file_uploader(
        "Upload Data File",
        type=["csv", "xlsx", "xls"],
        help="Upload CSV or Excel file with your data",
    )

    can_proceed = False

    if data_file:
        try:
            file_bytes = data_file.read()
            handler = DataHandler(file_bytes, data_file.name)
            st.session_state.data_handler = handler

            st.markdown(
                """
            <div class="success-box">
                <strong>✅ Data loaded successfully!</strong>
            </div>
            """,
                unsafe_allow_html=True,
            )

            col1, col2 = st.columns(2)
            with col1:
                st.metric("Total Rows", handler.get_row_count())
            with col2:
                st.metric("Total Columns", len(handler.get_columns()))

            # Data preview
            with st.expander("📋 Preview Data", expanded=True):
                st.dataframe(handler.get_preview(), use_container_width=True)

            st.session_state.step = max(st.session_state.step, 3)
            can_proceed = True

        except Exception as e:
            st.error(f"❌ Error loading data: {str(e)}")
            st.session_state.data_handler = None

    # Navigation buttons
    render_nav_buttons(1, can_proceed=can_proceed, next_label="Next: Map Columns ➡️")


def render_column_mapping():
    """Render column mapping section."""
    if (
        st.session_state.template_processor is None
        or st.session_state.data_handler is None
    ):
        st.info(
            "👆 Please complete the previous steps first. Click 'Back' to go to the previous step."
        )
        render_nav_buttons(2, can_proceed=False, show_next=False)
        return

    st.markdown("### 🔗 Step 3: Map Columns to Placeholders")

    placeholders = st.session_state.template_processor.get_placeholders()
    columns = st.session_state.data_handler.get_columns()

    # Auto-mapping logic
    if not st.session_state.column_mapping:
        auto_mapping = {}
        for placeholder in placeholders:
            # Try exact match first
            if placeholder in columns:
                auto_mapping[placeholder] = placeholder
            else:
                # Try case-insensitive match
                for col in columns:
                    if col.lower() == placeholder.lower():
                        auto_mapping[placeholder] = col
                        break
                    # Try partial match
                    elif (
                        placeholder.lower() in col.lower()
                        or col.lower() in placeholder.lower()
                    ):
                        auto_mapping[placeholder] = col
                        break
        st.session_state.column_mapping = auto_mapping

    st.markdown(
        """
    <div class="info-box">
        <strong>Map each placeholder to a column</strong> from your data file. 
        We've auto-detected some matches for you.
    </div>
    """,
        unsafe_allow_html=True,
    )

    # Create mapping UI
    mapping = {}
    columns_with_empty = ["-- Select Column --"] + columns

    col1, col2 = st.columns(2)

    for idx, placeholder in enumerate(placeholders):
        with col1 if idx % 2 == 0 else col2:
            current_value = st.session_state.column_mapping.get(placeholder, "")

            # Find index of current value
            try:
                default_idx = (
                    columns_with_empty.index(current_value) if current_value else 0
                )
            except ValueError:
                default_idx = 0

            selected = st.selectbox(
                f"📌 {{{placeholder}}}",
                options=columns_with_empty,
                index=default_idx,
                key=f"mapping_{placeholder}",
            )

            if selected and selected != "-- Select Column --":
                mapping[placeholder] = selected

    st.session_state.column_mapping = mapping

    # Validation
    unmapped = [p for p in placeholders if p not in mapping]
    can_proceed = False
    if unmapped:
        st.warning(
            f"⚠️ Unmapped placeholders: {', '.join(['{' + p + '}' for p in unmapped])}"
        )
    else:
        st.success("✅ All placeholders mapped!")
        st.session_state.step = max(st.session_state.step, 4)
        can_proceed = True

    # Filename pattern builder
    st.markdown("---")
    st.markdown("**📁 Output Filename Settings**")

    # Initialize filename pattern in session state
    if "filename_pattern" not in st.session_state:
        st.session_state.filename_pattern = []
    if "filename_mode" not in st.session_state:
        st.session_state.filename_mode = "auto"

    # Mode selection
    filename_mode = st.radio(
        "Filename Mode",
        options=["auto", "single", "pattern"],
        format_func=lambda x: {
            "auto": "🔢 Auto-numbered (document_001, document_002, ...)",
            "single": "📄 Single column",
            "pattern": "🎨 Custom pattern (multiple columns + separators)",
        }[x],
        horizontal=True,
        key="filename_mode_radio",
    )
    st.session_state.filename_mode = filename_mode

    if filename_mode == "single":
        # Simple single column selection
        filename_col = st.selectbox(
            "Select column for naming",
            options=columns,
            help="Choose a column value to use as the filename",
            key="single_filename_col",
        )
        st.session_state.filename_column = filename_col
        st.session_state.filename_pattern = []

        # Preview
        if st.session_state.data_handler:
            preview_data = st.session_state.data_handler.get_preview(1)
            if not preview_data.empty and filename_col in preview_data.columns:
                sample_value = str(preview_data[filename_col].iloc[0])
                st.markdown(f"**Preview:** `{sample_value}.docx`")

    elif filename_mode == "pattern":
        st.session_state.filename_column = None

        st.markdown(
            """
        <div class="info-box">
            <strong>Build your filename pattern:</strong> Add columns and separators in the order you want them to appear.
            <br><br>
            <strong>Example:</strong> <code>CustomerName</code> + <code>_</code> + <code>InvoiceNo</code> → <code>JohnDoe_INV001.docx</code>
        </div>
        """,
            unsafe_allow_html=True,
        )

        # Pattern builder
        col1, col2, col3 = st.columns([2, 1, 1])

        with col1:
            new_column = st.selectbox(
                "Add Column",
                options=["-- Select --"] + columns,
                key="pattern_add_column",
            )

        with col2:
            separator = st.text_input(
                "Separator",
                value="_",
                max_chars=5,
                help="Text between columns (e.g., _, -, ., space)",
                key="pattern_separator",
            )

        with col3:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("➕ Add to Pattern", key="add_to_pattern"):
                if new_column != "-- Select --":
                    # Add separator if pattern already has items
                    if st.session_state.filename_pattern:
                        st.session_state.filename_pattern.append(
                            {"type": "separator", "value": separator}
                        )
                    st.session_state.filename_pattern.append(
                        {"type": "column", "value": new_column}
                    )
                    st.rerun()

        # Display current pattern
        if st.session_state.filename_pattern:
            st.markdown("**Current Pattern:**")

            pattern_display = ""
            pattern_tags = []
            for item in st.session_state.filename_pattern:
                if item["type"] == "column":
                    pattern_tags.append(
                        f'<span class="placeholder-tag">{{{item["value"]}}}</span>'
                    )
                    pattern_display += f"{{{item['value']}}}"
                else:
                    pattern_tags.append(
                        f'<span style="color: #888; font-weight: bold;">{item["value"]}</span>'
                    )
                    pattern_display += item["value"]

            st.markdown(" ".join(pattern_tags), unsafe_allow_html=True)

            # Preview with actual data
            if st.session_state.data_handler:
                preview_data = st.session_state.data_handler.get_preview(1)
                if not preview_data.empty:
                    preview_filename = ""
                    for item in st.session_state.filename_pattern:
                        if (
                            item["type"] == "column"
                            and item["value"] in preview_data.columns
                        ):
                            preview_filename += str(preview_data[item["value"]].iloc[0])
                        elif item["type"] == "separator":
                            preview_filename += item["value"]
                    st.markdown(f"**Preview:** `{preview_filename}.docx`")

            # Clear pattern button
            col1, col2 = st.columns([1, 3])
            with col1:
                if st.button("🗑️ Clear Pattern", key="clear_pattern"):
                    st.session_state.filename_pattern = []
                    st.rerun()
        else:
            st.info("👆 Add columns and separators to build your filename pattern")

    else:  # auto mode
        st.session_state.filename_column = None
        st.session_state.filename_pattern = []
        st.markdown("**Preview:** `document_0001.docx`, `document_0002.docx`, ...")

    # Navigation buttons
    render_nav_buttons(2, can_proceed=can_proceed, next_label="Next: Generate ➡️")


def render_generate_section():
    """Render document generation section."""
    if st.session_state.step < 4:
        st.info(
            "👆 Please complete the column mapping first. Click 'Back' to go to the previous step."
        )
        render_nav_buttons(3, can_proceed=False, show_next=False)
        return

    st.markdown("### 🚀 Step 4: Generate Documents")

    processor = st.session_state.template_processor
    handler = st.session_state.data_handler
    mapping = st.session_state.column_mapping

    # Summary - determine filename mode description
    filename_mode = st.session_state.get("filename_mode", "auto")
    if filename_mode == "pattern" and st.session_state.get("filename_pattern"):
        pattern_desc = " + ".join(
            [
                f"{{{item['value']}}}"
                if item["type"] == "column"
                else f"'{item['value']}'"
                for item in st.session_state.filename_pattern
            ]
        )
        filename_desc = f"Pattern: {pattern_desc}"
    elif filename_mode == "single" and st.session_state.get("filename_column"):
        filename_desc = f"Column: {st.session_state.filename_column}"
    else:
        filename_desc = "Auto-numbered"

    st.markdown(
        f"""
    <div class="info-box">
        <strong>Summary:</strong>
        <ul>
            <li>📄 Documents to generate: <strong>{handler.get_row_count()}</strong></li>
            <li>🔗 Mapped placeholders: <strong>{len(mapping)}</strong></li>
            <li>📁 Filename: <strong>{filename_desc}</strong></li>
        </ul>
    </div>
    """,
        unsafe_allow_html=True,
    )

    # Preview section
    with st.expander("👁️ Preview First Document", expanded=False):
        data_rows = handler.get_data_as_dicts(mapping)
        if data_rows:
            first_row = data_rows[0]
            st.markdown("**Data for first document:**")
            preview_data = {f"{{{k}}}": v for k, v in first_row.items() if k in mapping}
            st.json(preview_data)

    # Generate button
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if st.button("🎯 Generate All Documents", use_container_width=True):
            with st.spinner("Generating documents..."):
                try:
                    data_rows = handler.get_data_as_dicts(mapping)

                    # Progress bar
                    progress_bar = st.progress(0)
                    status_text = st.empty()

                    documents = []
                    total = len(data_rows)

                    for idx, row_data in enumerate(data_rows):
                        # Generate filename based on mode
                        import re

                        filename_mode = st.session_state.get("filename_mode", "auto")

                        if filename_mode == "pattern" and st.session_state.get(
                            "filename_pattern"
                        ):
                            # Build filename from pattern
                            filename_parts = []
                            for item in st.session_state.filename_pattern:
                                if (
                                    item["type"] == "column"
                                    and item["value"] in row_data
                                ):
                                    filename_parts.append(str(row_data[item["value"]]))
                                elif item["type"] == "separator":
                                    filename_parts.append(item["value"])
                            filename = "".join(filename_parts) + ".docx"
                            filename = re.sub(r'[<>:"/\\|?*]', "_", filename)
                        elif filename_mode == "single" and st.session_state.get(
                            "filename_column"
                        ):
                            if st.session_state.filename_column in row_data:
                                filename = (
                                    f"{row_data[st.session_state.filename_column]}.docx"
                                )
                                filename = re.sub(r'[<>:"/\\|?*]', "_", filename)
                            else:
                                filename = f"document_{idx + 1:04d}.docx"
                        else:
                            # Auto mode
                            filename = f"document_{idx + 1:04d}.docx"

                        # Generate document
                        doc_bytes = processor.generate_document(row_data)
                        documents.append((filename, doc_bytes))

                        # Update progress
                        progress = (idx + 1) / total
                        progress_bar.progress(progress)
                        status_text.text(f"Generated: {idx + 1}/{total} documents")

                    st.session_state.generated_docs = documents
                    status_text.empty()
                    progress_bar.empty()

                    st.success(f"✅ Successfully generated {len(documents)} documents!")

                except Exception as e:
                    st.error(f"❌ Error generating documents: {str(e)}")

    # Download section
    if st.session_state.generated_docs:
        st.markdown("---")
        st.markdown("### 📥 Download Generated Documents")

        documents = st.session_state.generated_docs

        col1, col2 = st.columns(2)

        with col1:
            # Download as ZIP
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for filename, doc_bytes in documents:
                    zip_file.writestr(filename, doc_bytes)

            st.download_button(
                label="📦 Download All as ZIP",
                data=zip_buffer.getvalue(),
                file_name="generated_documents.zip",
                mime="application/zip",
                use_container_width=True,
            )

        with col2:
            st.metric("Total Documents", len(documents))

        # Individual downloads
        with st.expander("📄 Download Individual Documents"):
            for filename, doc_bytes in documents[:20]:  # Show first 20
                st.download_button(
                    label=f"📄 {filename}",
                    data=doc_bytes,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key=f"download_{filename}",
                )

            if len(documents) > 20:
                st.info(
                    f"Showing first 20 of {len(documents)} documents. Download ZIP for all."
                )

    # Navigation buttons
    can_proceed_to_email = st.session_state.generated_docs is not None
    render_nav_buttons(
        3, can_proceed=can_proceed_to_email, next_label="Next: Send Emails ➡️"
    )


def render_email_section():
    """Render the Email/Send section (Step 5)."""
    st.header("📤 Send Documents")
    
    if "generated_docs" not in st.session_state or not st.session_state.generated_docs:
        st.warning("⚠️ No documents generated yet. Please complete the Generate step first.")
        render_nav_buttons(4, can_proceed=False, show_next=False)
        return

    # Toggle between Email and DocuSign
    send_mode = st.radio(
        "Select Action:", 
        ["📧 Send via Email (SMTP)", "✍️ Send for Signature (DocuSign)"], 
        horizontal=True,
        help="Choose 'Email' to send documents as attachments. Choose 'DocuSign' to request e-signatures."
    )

    if send_mode == "📧 Send via Email (SMTP)":
        render_smtp_email_section()
    else:
        render_docusign_logic()
    
    # Navigation buttons (only if not in DocuSign mode, as that has its own flow/buttons usually)
    # But for consistency we can keep them at bottom or let the sub-functions handle it.
    # Current implementation of sub-functions ends with content, no nav.
    render_nav_buttons(4, can_proceed=False, show_next=False)


def render_smtp_email_section():
    """Render the SMTP Email sending interface."""
    handler = st.session_state.data_handler
    data_df = handler.df
    generated_docs = st.session_state.generated_docs
    
    # Initialize session state for email handler
    if "email_handler" not in st.session_state:
        st.session_state.email_handler = None
    if "email_configured" not in st.session_state:
        st.session_state.email_configured = False
    
    # Configuration Section (Expander)
    with st.expander("⚙️ Email Configuration", expanded=not st.session_state.email_configured):
        col1, col2 = st.columns(2)
        with col1:
            smtp_server = st.text_input("SMTP Server", value=st.session_state.get("smtp_server", "smtp.gmail.com"))
            smtp_port = st.number_input("SMTP Port", value=st.session_state.get("smtp_port", 587))
            sender_email = st.text_input("Sender Email", value=st.session_state.get("sender_email", ""))
        with col2:
            sender_name = st.text_input("Sender Name", value=st.session_state.get("sender_name", ""))
            sender_password = st.text_input("App Password", type="password", help="Use App Password for Gmail", value=st.session_state.get("sender_password", ""))
        
        # Store values in session state
        st.session_state.smtp_server = smtp_server
        st.session_state.smtp_port = smtp_port
        st.session_state.sender_email = sender_email
        st.session_state.sender_name = sender_name
        st.session_state.sender_password = sender_password

        if st.button("Connect & Verify"):
            if not (sender_email and sender_password):
                st.error("Please provide email and password")
            else:
                handler_obj = EmailHandler(smtp_server, smtp_port, sender_email, sender_password, sender_name)
                success, msg = handler_obj.test_connection() # Using test_connection derived from previous file read
                if success:
                    st.success(msg)
                    st.session_state.email_handler = handler_obj
                    st.session_state.email_configured = True
                else:
                    st.error(msg)
    
    if st.session_state.email_configured:
        st.markdown("---")
        st.success(f"✅ Connected as: {st.session_state.email_handler.sender_email}")
        
        # Email Column Selection Logic
        columns = handler.get_columns()
        email_column_suggestions = [col for col in columns if 'email' in col.lower() or 'e-mail' in col.lower() or 'mail' in col.lower()]
        default_email_col = email_column_suggestions[0] if email_column_suggestions else (columns[0] if columns else None)
        
        if "email_column" not in st.session_state or st.session_state.email_column is None:
             st.session_state.email_column = default_email_col
             
        email_column = st.selectbox(
            "Column containing recipient email addresses",
            options=columns,
            index=columns.index(st.session_state.email_column) if st.session_state.email_column in columns else 0,
            key="email_column_select_smtp",
        )
        st.session_state.email_column = email_column
        
        # Run Validation Logic
        validation_results = {
            "valid": [],
            "missing": [],
            "invalid": [],
        }
        
        # Initialize missing_emails and skip_rows if not present
        if "missing_emails" not in st.session_state:
            st.session_state.missing_emails = {}
        if "skip_rows" not in st.session_state:
            st.session_state.skip_rows = set()

        for idx, row in data_df.iterrows():
            email_value = row.get(email_column, "")
            
            if not email_value or str(email_value).strip() == "" or str(email_value).lower() == "nan":
                validation_results["missing"].append({
                    "row_index": idx,
                    "row_data": row.to_dict(),
                })
            elif EmailHandler.validate_email(str(email_value)):
                validation_results["valid"].append({
                    "row_index": idx,
                    "email": str(email_value).strip(),
                })
            else:
                validation_results["invalid"].append({
                    "row_index": idx,
                    "email": str(email_value),
                    "row_data": row.to_dict(),
                })
        
        results = validation_results # Local alias

        # Preview Section (Recipients)
        total_valid = len(results["valid"])
        total_missing = len(results["missing"])
        total_invalid = len(results["invalid"])
        
        st.markdown(f"**Target Recipients:** {total_valid} valid emails found.")
        
        # Calculate total sendable (valid + manually fixed - skipped)
        manual_emails_count = len([e for k, e in st.session_state.missing_emails.items() if k not in st.session_state.skip_rows and EmailHandler.validate_email(e)])
        valid_emails_count = len([r for r in results["valid"] if r["row_index"] not in st.session_state.skip_rows])
        total_sendable = valid_emails_count + manual_emails_count
        total_skipped = len(st.session_state.skip_rows)
        
        st.metric("Total Emails to Send", total_sendable, delta=f"-{total_skipped} skipped", delta_color="off")

        # Missing Emails Handling
        if total_missing > 0:
            with st.expander(f"⚠️ Missing Emails ({total_missing})", expanded=True):
                st.warning("The following rows are missing email addresses. Provide them manually or skip.")
                
                display_columns = list(data_df.columns)[:3] 
                
                for item in results["missing"][:20]: # Limit to first 20 for UI perf
                    row_idx = item["row_index"]
                    row_data = item["row_data"]
                    
                    col1, col2, col3 = st.columns([2, 2, 1])
                    
                    with col1:
                        display_info = " | ".join([f"{col}: {row_data.get(col, 'N/A')}" for col in display_columns])
                        st.markdown(f"**Row {row_idx + 1}:** {display_info}")
                    
                    with col2:
                        manual_email = st.text_input(
                            "Enter email",
                            key=f"missing_email_{row_idx}",
                            placeholder="email@example.com",
                            label_visibility="collapsed",
                        )
                        if manual_email:
                            st.session_state.missing_emails[row_idx] = manual_email
                    
                    with col3:
                        skip = st.checkbox("Skip", key=f"skip_{row_idx}", value=row_idx in st.session_state.skip_rows)
                        if skip:
                            st.session_state.skip_rows.add(row_idx)
                        elif row_idx in st.session_state.skip_rows:
                            st.session_state.skip_rows.remove(row_idx)
                
                if len(results["missing"]) > 20:
                    st.info(f"Showing first 20 of {len(results['missing'])} missing emails")

        # Invalid emails section
        if results["invalid"]:
            st.markdown("---")
            st.markdown("##### ❌ Invalid Email Formats")
            
            for invalid_item in results["invalid"][:10]:  # Show first 10
                row_idx = invalid_item["row_index"]
                current_email = invalid_item["email"]
                row_data = invalid_item["row_data"]
                
                col1, col2 = st.columns([1, 2])
                
                with col1:
                    st.markdown(f"**Row {row_idx + 1}:** `{current_email}`")
                
                with col2:
                    corrected_email = st.text_input(
                        "Correct email",
                        key=f"invalid_email_{row_idx}",
                        value=current_email,
                        label_visibility="collapsed",
                    )
                    if corrected_email != current_email:
                        st.session_state.missing_emails[row_idx] = corrected_email

        # Email template section
        st.markdown("---")
        st.markdown("#### 3️⃣ Email Template")
        
        # Show available placeholders
        mapping = st.session_state.column_mapping
        available_placeholders = list(mapping.keys())
        
        st.markdown(
            f"""
            <div class="info-box">
                <strong>Available Placeholders:</strong> {", ".join([f"<code>{{{p}}}</code>" for p in available_placeholders])}
            </div>
            """,
            unsafe_allow_html=True,
        )
        
        # Email recipients section (CC/BCC)
        col1, col2 = st.columns(2)
        with col1:
            cc_emails = st.text_input(
                "CC (Optional)",
                value=st.session_state.get("cc_emails", ""),
                placeholder="email1@example.com, email2@example.com",
                help="Separate multiple emails with commas",
                key="cc_emails_input"
            )
            st.session_state.cc_emails = cc_emails
        
        with col2:
            bcc_emails = st.text_input(
                "BCC (Optional)",
                value=st.session_state.get("bcc_emails", ""),
                placeholder="email1@example.com, email2@example.com",
                help="Separate multiple emails with commas",
                key="bcc_emails_input"
            )
            st.session_state.bcc_emails = bcc_emails
        
        # Set defaults if empty
        if "email_subject_template" not in st.session_state or not st.session_state.email_subject_template:
            st.session_state.email_subject_template = "Document for {" + (available_placeholders[0] if available_placeholders else "Recipient") + "}"
        
        if "email_body_template" not in st.session_state or not st.session_state.email_body_template:
            st.session_state.email_body_template = "Dear Recipient,\n\nPlease find your document attached.\n\nBest regards"
        
        st.text_input("Email Subject", key="email_subject_template", help="Use {placeholder} format")
        st.text_area("Email Body", key="email_body_template", height=200, help="Use {placeholder} format")
        st.caption("ℹ️ Changes are applied automatically when you click outside the text box.")
        
        # Batch Attachment Section
        st.markdown("**📎 Batch Attachments (Map files to rows)**")
        
        # Initializing session state for batch attachments (reusing from previous logic)
        if "batch_attachments" not in st.session_state:
            st.session_state.batch_attachments = {}
        if "batch_mapping_column" not in st.session_state:
            st.session_state.batch_mapping_column = None
        if "batch_prefix" not in st.session_state:
            st.session_state.batch_prefix = ""

        col1, col2 = st.columns([2, 1])
        
        with col1:
            batch_files = st.file_uploader(
                "Upload Batch Files",
                type=["pdf", "doc", "docx", "xls", "xlsx", "jpg", "png"],
                accept_multiple_files=True,
                help="Upload all files you want to attach.",
                key="batch_files_uploader"
            )

        with col2:
            mapping_cols = ["-- Select Column --"] + data_df.columns.tolist()
            batch_col = st.selectbox(
                "Match Filenames with Column",
                options=mapping_cols,
                help="Select column matching filenames",
                key="batch_mapping_col_select",
                index=mapping_cols.index(st.session_state.batch_mapping_column) if st.session_state.batch_mapping_column in mapping_cols else 0
            )
            st.session_state.batch_mapping_column = batch_col
            
            batch_prefix = st.text_input(
                "Filename Prefix",
                value=st.session_state.batch_prefix,
                help="Prefix expected in filenames",
                key="batch_prefix_input"
            )
            st.session_state.batch_prefix = batch_prefix

        # Process batch upload (simplified reuse of efficient logic)
        if batch_files and batch_col != "-- Select Column --":
            # ... process batch files logic ...
            mapping_dict = {}
            for idx, row in data_df.iterrows():
                 val = str(row[batch_col]).strip().lower()
                 if val: mapping_dict[val] = idx
            
            matched_count = 0
            # Reset
            st.session_state.batch_attachments = {}

            for uploaded_file in batch_files:
                fname = uploaded_file.name
                clean_name = fname.rsplit(".", 1)[0].lower()
                
                if batch_prefix and batch_prefix.lower() in clean_name:
                     # Remove prefix logic if needed, but usually matching the value is enough if prefix is consistent
                     # For now assuming value is part of filename
                     pass
                
                # Try logic: if filename (no ext) ends with the value? Or equals?
                # Simple logic: check if any key in mapping_dict matches filename
                # Better: Check if clean_name contains the ID
                
                # Re-implementing specific logical match from previous step:
                # 1. Strip prefix from filename
                matched_row_idx = None
                
                # Check direct match after prefix removal
                candidate_name = clean_name
                if batch_prefix and batch_prefix.lower() in candidate_name:
                     if candidate_name.startswith(batch_prefix.lower()):
                         candidate_name = candidate_name[len(batch_prefix):]
                
                if candidate_name in mapping_dict:
                     matched_row_idx = mapping_dict[candidate_name]
                
                if matched_row_idx is not None:
                     if matched_row_idx not in st.session_state.batch_attachments:
                         st.session_state.batch_attachments[matched_row_idx] = []
                     st.session_state.batch_attachments[matched_row_idx].append((uploaded_file.name, uploaded_file.getvalue()))
                     matched_count += 1
            
            if matched_count > 0:
                st.success(f"✅ Mapped {matched_count} files to {len(st.session_state.batch_attachments)} rows!")

        # Common Attachment
        st.markdown("**📄 Common Attachment (Optional)**")
        col1, col2 = st.columns([2, 1])
        with col1:
            if "common_attachment" not in st.session_state:
                st.session_state.common_attachment = None
            
            common_file = st.file_uploader("Upload One File for Everyone", type=["pdf", "doc", "docx", "jpg", "png", "zip"], key="common_file_uploader")
            if common_file:
                st.session_state.common_attachment = (common_file.name, common_file.getvalue())
                st.success(f"✅ Attached: {common_file.name}")
        with col2:
             if st.session_state.common_attachment:
                 st.markdown("<br>", unsafe_allow_html=True)
                 if st.button("🗑️ Remove", use_container_width=True):
                     st.session_state.common_attachment = None
                     st.rerun()

        # Build Preview Data
        preview_data = []
        if total_sendable > 0:
             # Just show first valid one
             first_valid = next((r for r in results["valid"] if r["row_index"] not in st.session_state.skip_rows), None)
             if first_valid:
                 row_idx = first_valid["row_index"]
                 row_data = data_df.iloc[row_idx].to_dict()
                 display_info = " | ".join([f"{col}: {row_data.get(col, 'N/A')}" for col in list(data_df.columns)[:3]])
                 
                 preview_data.append({
                    "Row": row_idx + 1,
                    "Recipient Info": display_info,
                    "Email": first_valid["email"],
                    "Document": generated_docs[row_idx][0],
                    "Batch File": st.session_state.batch_attachments[row_idx][0][0] if "batch_attachments" in st.session_state and row_idx in st.session_state.batch_attachments else "No",
                })
        
        if preview_data:
            with st.expander("👁️ Preview First Email", expanded=False):
                st.table(preview_data)
                
                # Render content preview
                idx = preview_data[0]["Row"] - 1
                row_data_prev = data_df.iloc[idx].to_dict()
                subject_prev = EmailHandler.render_template(st.session_state.email_subject_template, row_data_prev)
                body_prev = EmailHandler.render_template(st.session_state.email_body_template, row_data_prev)
                
                st.markdown(f"**Subject:** {subject_prev}")
                st.text_area("Body Preview", value=body_prev, disabled=True, height=150)


        # Send Logic
        st.markdown("---")
        st.markdown("#### 4️⃣ Send Emails")
        
        # Speed control
        col_a, col_b = st.columns([3, 1])
        with col_a:
            delay_seconds = st.slider("Delay between emails (seconds)", 0.1, 2.0, 0.5, 0.1, help="Gmail limit: ~500/day")
        with col_b:
            estimated_time = total_sendable * delay_seconds
            st.metric("Est. Time", f"{estimated_time:.0f}s")
            
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            if "email_send_results" not in st.session_state:
                st.session_state.email_send_results = None

            if not st.session_state.email_send_results:
                if st.button("✅ Confirm & Send All Emails", type="primary", use_container_width=True):
                    # Prepare data
                    email_data_list = []
                    
                    # Add valid
                    for valid_item in results["valid"]:
                        row_idx = valid_item["row_index"]
                        if row_idx in st.session_state.skip_rows: continue
                        row_data = data_df.iloc[row_idx].to_dict()
                        
                        # Attachments logic
                        atts = []
                        if st.session_state.common_attachment: atts.append(st.session_state.common_attachment)
                        if "batch_attachments" in st.session_state and row_idx in st.session_state.batch_attachments:
                            # batch_attachments stores list of tuples (name, bytes), append them
                            for batch_file in st.session_state.batch_attachments[row_idx]:
                                atts.append(batch_file)
                                
                        email_data_list.append({
                            "to_email": valid_item["email"],
                            "subject": EmailHandler.render_template(st.session_state.email_subject_template, row_data),
                            "body": EmailHandler.render_template(st.session_state.email_body_template, row_data),
                            "attachment_filename": generated_docs[row_idx][0],
                            "attachment_data": generated_docs[row_idx][1],
                            "cc_emails": [e.strip() for e in st.session_state.cc_emails.split(',') if e.strip()] if st.session_state.cc_emails else None,
                            "bcc_emails": [e.strip() for e in st.session_state.bcc_emails.split(',') if e.strip()] if st.session_state.bcc_emails else None,
                            "additional_attachments": atts if atts else None,
                            "row_index": row_idx,
                            "recipient_info": " | ".join([f"{col}: {row_data.get(col, 'N/A')}" for col in list(data_df.columns)[:3]]),
                        })
                        
                    # Add manuals
                    for row_idx, manual_email in st.session_state.missing_emails.items():
                        if row_idx in st.session_state.skip_rows: continue
                        if not EmailHandler.validate_email(manual_email): continue
                        row_data = data_df.iloc[row_idx].to_dict()
                        
                        atts = []
                        if st.session_state.common_attachment: atts.append(st.session_state.common_attachment)
                        if "batch_attachments" in st.session_state and row_idx in st.session_state.batch_attachments:
                            for batch_file in st.session_state.batch_attachments[row_idx]:
                                atts.append(batch_file)

                        email_data_list.append({
                            "to_email": manual_email,
                            "subject": EmailHandler.render_template(st.session_state.email_subject_template, row_data),
                            "body": EmailHandler.render_template(st.session_state.email_body_template, row_data),
                            "attachment_filename": generated_docs[row_idx][0],
                            "attachment_data": generated_docs[row_idx][1],
                            "cc_emails": [e.strip() for e in st.session_state.cc_emails.split(',') if e.strip()] if st.session_state.cc_emails else None,
                            "bcc_emails": [e.strip() for e in st.session_state.bcc_emails.split(',') if e.strip()] if st.session_state.bcc_emails else None,
                            "additional_attachments": atts if atts else None,
                            "row_index": row_idx,
                            "recipient_info": " | ".join([f"{col}: {row_data.get(col, 'N/A')}" for col in list(data_df.columns)[:3]]),
                        })

                    # Send loop
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    def progress_callback(current, total, status):
                        progress_bar.progress(current / total)
                        status_text.text(f"{status} ({current}/{total})")
                    
                    send_results = st.session_state.email_handler.send_batch_emails(
                        email_data_list,
                        progress_callback=progress_callback,
                        delay_seconds=delay_seconds
                    )
                    
                    progress_bar.empty()
                    status_text.empty()
                    
                    # Store Results
                    detailed_results = []
                    for email_data in email_data_list:
                        row_idx = email_data["row_index"]
                        failed_item = next((f for f in send_results["failed_details"] if f["row_index"] == row_idx), None)
                        
                        detailed_results.append({
                            "Row": row_idx + 1,
                            "Recipient Info": email_data["recipient_info"],
                            "Email": email_data["to_email"],
                            "Document": email_data["attachment_filename"],
                            "Status": "❌ Failed" if failed_item else "✅ Sent",
                            "Error": failed_item["error"] if failed_item else "",
                        })
                    
                    st.session_state.email_send_results = {
                        "summary": send_results,
                        "details": detailed_results,
                    }
                    st.rerun()

        # Show Results (Shared UI)
        if st.session_state.email_send_results:
            results_data = st.session_state.email_send_results
            summary = results_data["summary"]
            
            # Summary metrics
            st.markdown("**📊 Summary:**")
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("✅ Sent Successfully", summary["sent"], delta=None, delta_color="normal")
            with col2:
                st.metric("❌ Failed", summary["failed"], delta=None, delta_color="inverse")
            with col3:
                st.metric("⏭️ Skipped", total_skipped)
            
            # Detailed results table
            st.markdown("**📋 Detailed Results:**")
            import pandas as pd
            results_df = pd.DataFrame(results_data["details"])
            st.dataframe(results_df, use_container_width=True, hide_index=True)






def render_sidebar():
    """Render the sidebar with instructions and info."""
    with st.sidebar:
        st.markdown("## 📚 Quick Guide")

        st.markdown("""
        ### How to Use
        
        1. **Create your template** in Word with placeholders:
           - Use `{Column Name}` format
           - E.g., `{Customer Name}`, `{Order ID}`
        
        2. **Prepare your data** in CSV/Excel:
           - First row should be headers
           - Headers should match placeholders
        
        3. **Upload both files** and map columns
        
        4. **Generate** and download your documents!
        
        ---
        
        ### Placeholder Examples
        
        ```
        Dear {Name},
        
        Your order #{Order ID} dated 
        {Order Date} is confirmed.
        
        Total: ₹{Amount}
        ```
        
        ---
        
        ### Tips 💡
        
        - Column names are **case-insensitive**
        - Placeholders work in **tables** too
        - Headers & footers are supported
        - Use meaningful column for filenames
        - **Number to Words**: Add `_Words` to any numeric column name to convert it to text (e.g., `{Amount}` → `{Amount_Words}`)
        """)

        st.markdown("---")
        st.markdown("### 🔄 Reset")
        if st.button("Clear All & Start Over"):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()


def main():
    """Main application entry point."""
    init_session_state()
    render_header()
    render_sidebar()


    # Tab names for navigation
    # Consolidating into "Send" tab
    tab_names = ["📝 Template", "📊 Data", "🔗 Mapping", "🚀 Generate", " Send"]

    # Create clickable step indicators - all as buttons for consistent UI
    cols = st.columns(5)
    for i, (col, name) in enumerate(zip(cols, tab_names)):
        with col:
            # Determine button style based on current tab
            is_current = st.session_state.current_tab == i
            is_completed = st.session_state.step > i + 1

            if is_current:
                # Current tab - show as active button (clicking does nothing as we're already here)
                st.button(
                    f"● {name}",
                    key=f"tab_btn_{i}",
                    use_container_width=True,
                    type="primary",
                )
            elif is_completed:
                if st.button(
                    f"✅ {name}", key=f"tab_btn_{i}", use_container_width=True
                ):
                    go_to_tab(i)
                    st.rerun()
            else:
                if st.button(name, key=f"tab_btn_{i}", use_container_width=True):
                    go_to_tab(i)
                    st.rerun()

    st.markdown("---")

    # Render current tab content
    current_tab = st.session_state.current_tab

    if current_tab == 0:
        render_template_upload()
    elif current_tab == 1:
        render_data_upload()
    elif current_tab == 2:
        render_column_mapping()
    elif current_tab == 3:
        render_generate_section()
    elif current_tab == 4:
        render_email_section()
    elif current_tab == 5:
        render_docusign_logic()




from utils.docusign_handler import DocuSignHandler
import os

def render_docusign_logic():
    """Render the DocuSign specific logic."""
    st.markdown("### ✍️ DocuSign Integration")
    
    if "docusign_results" not in st.session_state:
        st.session_state.docusign_results = None

    # Credentials Section
    # Use absolute path relative to app.py
    key_file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "docusign_key.txt")
    has_key_file = os.path.exists(key_file_path)
    
    # Auto-parse credentials if file exists and not yet set
    default_ik = ""
    default_uid = ""
    default_aid = ""
    default_base = "https://demo.docusign.net"

    if has_key_file:
        try:
            with open(key_file_path, "r") as f:
                content = f.read()
                for line in content.splitlines():
                    if "Integration Key =" in line: default_ik = line.split("=")[1].strip()
                    if "User ID =" in line: default_uid = line.split("=")[1].strip()
                    if "API Account ID =" in line: default_aid = line.split("=")[1].strip()
                    # if "Account Base URI =" in line: default_base = line.split("=")[1].strip() 
                    # Note: We intentionally IGNORE the file's base URI because it often defaults to Prod (na4)
                    # even for Developer keys. We stick to demo unless user changes it.
        except:
            pass
            
    with st.expander("🔑 DocuSign Credentials", expanded=True):
        if has_key_file:
            st.success("✅ 'docusign_key.txt' found and parsed.")
        else:
            st.error("❌ 'docusign_key.txt' missing. Please create it with your Private Key.")
            
        col1, col2 = st.columns(2)
        with col1:
            st.session_state.ds_integration_key = st.text_input("Integration Key", value=st.session_state.get("ds_integration_key", default_ik))
            st.session_state.ds_user_id = st.text_input("User ID", value=st.session_state.get("ds_user_id", default_uid))
        with col2:
            st.session_state.ds_account_id = st.text_input("API Account ID", value=st.session_state.get("ds_account_id", default_aid))
            
            # Smart default: If session has nothing, use demo. 
            current_base = st.session_state.get("ds_base_url", default_base)
            st.session_state.ds_base_url = st.text_input("Base URL", value=current_base, help="Use 'https://demo.docusign.net' for Developer Sandbox accounts.")
            
            if "na" in st.session_state.ds_base_url and "demo" not in st.session_state.ds_base_url:
                st.warning("⚠️ You are using a Production URL. New integrations usually require 'https://demo.docusign.net'.")

    # SMTP Configuration Section
    st.markdown("---")
    with st.expander("📧 SMTP Email Configuration", expanded=True):
        st.info("DocuSign will generate signing links and send them via your email server.")
        col1, col2 = st.columns(2)
        with col1:
            smtp_server = st.text_input("SMTP Server", value=st.session_state.get("smtp_server", "smtp.gmail.com"))
            smtp_port = st.number_input("SMTP Port", value=st.session_state.get("smtp_port", 587), min_value=1, max_value=65535)
            sender_email = st.text_input("Sender Email", value=st.session_state.get("sender_email", ""))
        with col2:
            sender_name = st.text_input("Sender Name", value=st.session_state.get("sender_name", ""))
            sender_password = st.text_input("App Password", type="password", help="Use App Password for Gmail", value=st.session_state.get("sender_password", ""))
        
        # Store values in session state
        st.session_state.smtp_server = smtp_server
        st.session_state.smtp_port = smtp_port
        st.session_state.sender_email = sender_email
        st.session_state.sender_name = sender_name
        st.session_state.sender_password = sender_password
        
        # Test connection button
        if st.button("🔌 Connect & Verify SMTP", key="test_smtp_ds"):
            if not (sender_email and sender_password):
                st.error("❌ Please provide email and password")
            else:
                try:
                    from utils.email_handler import EmailHandler
                    handler_obj = EmailHandler(smtp_server, smtp_port, sender_email, sender_password, sender_name)
                    success, msg = handler_obj.test_connection()
                    if success:
                        st.success(f"✅ {msg}")
                        st.session_state.email_handler = handler_obj
                        st.session_state.email_configured = True
                    else:
                        st.error(f"❌ {msg}")
                except Exception as e:
                    st.error(f"❌ Connection failed: {str(e)}")

    # Batch File Mapping Section (moved below email body, expanded by default)
    st.markdown("---")
    st.markdown("### 📎 Batch File Attachments")
    st.info("Upload files that will be attached to specific recipients based on filename matching. Each recipient gets only their matched file(s).")
    
    if True:  # Always show, not in expander
        if "data_handler" in st.session_state and st.session_state.data_handler is not None:
            data_df = st.session_state.data_handler.df
            available_columns = data_df.columns.tolist()
            
            # Batch file configuration
            batch_files = st.file_uploader("Upload Batch Files", accept_multiple_files=True, key="ds_batch_files", 
                                          help="Files will be matched to recipients based on the filename column you select below")
            
            if batch_files:
                col1, col2 = st.columns(2)
                with col1:
                    batch_filename_col = st.selectbox("Filename Column (for matching)", options=available_columns, 
                                                     help="Column containing values that match your batch filenames")
                with col2:
                    batch_prefix = st.text_input("Filename Prefix (optional)", placeholder="e.g., 'invoice_'",
                                                help="If your files have a prefix like 'invoice_001.pdf', enter 'invoice_'")
                
                if st.button("🔗 Map Batch Files", key="map_batch_ds"):
                    # Initialize batch_attachments if not present
                    if "batch_attachments" not in st.session_state:
                        st.session_state.batch_attachments = {}
                    
                    # Create mapping dictionary
                    mapping_dict = {}
                    for idx, row in data_df.iterrows():
                        key_value = str(row[batch_filename_col]).strip().lower()
                        mapping_dict[key_value] = idx
                    
                    # Match files
                    matched_count = 0
                    for uploaded_file in batch_files:
                        clean_name = uploaded_file.name.lower().replace(" ", "").split(".")[0]
                        
                        # Strip prefix if provided
                        candidate_name = clean_name
                        if batch_prefix and batch_prefix.lower() in candidate_name:
                            if candidate_name.startswith(batch_prefix.lower()):
                                candidate_name = candidate_name[len(batch_prefix):]
                        
                        if candidate_name in mapping_dict:
                            matched_row_idx = mapping_dict[candidate_name]
                            if matched_row_idx not in st.session_state.batch_attachments:
                                st.session_state.batch_attachments[matched_row_idx] = []
                            st.session_state.batch_attachments[matched_row_idx].append((uploaded_file.name, uploaded_file.getvalue()))
                            matched_count += 1
                    
                    if matched_count > 0:
                        st.success(f"✅ Mapped {matched_count} files to {len(st.session_state.batch_attachments)} recipients!")
                    else:
                        st.warning("⚠️ No files matched. Check your filename column and prefix settings.")
        else:
            st.warning("Please upload data in the Data step first.")

    # Mapping Section
    st.markdown("### 🔗 Recipient Mapping & Email Config")
    st.markdown("_This uses your Email Configuration (SMTP) to send the signing link._")
    
    if "data_handler" not in st.session_state or st.session_state.data_handler is None:
         st.warning("Please upload data first.")
         return

    data_df = st.session_state.data_handler.df
    available_columns = data_df.columns.tolist()
    
    col1, col2 = st.columns(2)
    with col1:
        recipient_email_col = st.selectbox("Recipient Email Column", options=available_columns, key="ds_email_col", index=available_columns.index("Email") if "Email" in available_columns else 0)
    with col2:
        recipient_name_col = st.selectbox("Recipient Name Column", options=available_columns, key="ds_name_col", index=available_columns.index("Name") if "Name" in available_columns else 0)

    # CC/BCC (moved above subject)
    col_em1, col_em2 = st.columns(2)
    with col_em1:
        ds_cc_emails = st.text_input("CC (Optional)", placeholder="email1@example.com, email2@example.com")
    with col_em2:
        ds_bcc_emails = st.text_input("BCC (Optional)", placeholder="email1@example.com, email2@example.com")
    
    # Email Subject/Body
    ds_email_subject = st.text_input("Email Subject", value="Action Required: Please Sign Document {Filename}", help="Placeholders: {Filename}, {Name}")
    ds_email_body = st.text_area("Email Body", value="Dear {Name},\n\nPlease review and sign the attached document by clicking the link below:\n\n{Signing_Link}\n\nBest regards,", height=150)
    st.caption("ℹ️ The `{Signing_Link}` placeholder will be replaced by the unique DocuSign link.")
    
    # Common attachments
    st.markdown("**📂 Common Attachments**")
    ds_additional_files = st.file_uploader("Attach extra files to all emails", accept_multiple_files=True, help="These files will be attached to every email")

    # Sending Logic
    st.markdown("---")
    
    # Delivery Method Selection
    st.markdown("#### 📨 Delivery Method")
    delivery_method = st.radio(
        "Choose how the signing request is sent:",
        ["DocuSign Official Email (Recommended - Never Expires)", "My SMTP Email (Link expires in 5 minutes!)"],
        help="DocuSign's email service provides secure, non-expiring links. Sending via your own email requires generating a temporary link that expires quickly for security."
    )
    
    use_docusign_email = "DocuSign" in delivery_method
    
    if st.session_state.docusign_results:
        # Show Results
        results = st.session_state.docusign_results
        col1, col2 = st.columns(2)
        with col1: st.metric("✅ Sent", results["sent"])
        with col2: st.metric("❌ Failed", results["failed"])
             
        if results["details"]:
            st.dataframe(results["details"], use_container_width=True)
            
        if st.button("🔄 Send New Batch", key="reset_ds"):
            st.session_state.docusign_results = None
            st.rerun()
            
    else:
        # Send Button 
        if st.button("🚀 Generate & Send", type="primary", use_container_width=True):
            if not has_key_file:
                st.error("Missing Private Key file.")
                return

            # Check SMTP configuration (Only strictly needed if NOT using DocuSign email, or if we want to send notification)
            # But kept for consistency/logging
            if not use_docusign_email and (not st.session_state.get("email_configured") or not st.session_state.get("email_handler")):
                st.error("⚠️ Please configure and connect SMTP Email first (in the 'Send via Email' tab) to send links yourself.")
                return

            # Init DocuSign
            try:
                ds_integration_key = st.session_state.get("ds_integration_key", "f588bd38-c3ac-428b-a6be-b9efbde23a5a")
                ds_user_id = st.session_state.get("ds_user_id", "c1370aa2-d90e-493e-a329-b2ef724105a5")
                ds_account_id = st.session_state.get("ds_account_id", "516db58d-941b-4954-a0b1-19c9723f07fb")
                ds_base_url = st.session_state.get("ds_base_url", "https://demo.docusign.net")
                
                ds_handler = DocuSignHandler(
                    ds_integration_key,
                    ds_user_id,
                    ds_account_id,
                    key_file_path, 
                    ds_base_url
                )
            except Exception as e:
                st.error(f"DocuSign Init Error: {str(e)}")
                return
            
            st.success("✅ Connected to DocuSign! Processing...")
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            results = {"sent": 0, "failed": 0, "details": []}
            generated_docs = st.session_state.generated_docs
            total_docs = len(generated_docs)
            
            # Prepare common attachments (Read once)
            common_attachments_data = []
            if ds_additional_files:
                for f in ds_additional_files:
                    f.seek(0)
                    common_attachments_data.append((f.name, f.read()))

            # Parse CC/BCC
            cc_list = [e.strip() for e in ds_cc_emails.split(",")] if ds_cc_emails else []
            bcc_list = [e.strip() for e in ds_bcc_emails.split(",")] if ds_bcc_emails else [] # BCC only works for SMTP

            for i, (filename, file_data) in enumerate(generated_docs):
                progress_bar.progress((i + 1) / total_docs)
                status_text.text(f"Processing {i+1}/{total_docs}: {filename}...")
                
                if i < len(data_df):
                    row = data_df.iloc[i]
                    rec_email = str(row[recipient_email_col]).strip()
                    rec_name = str(row[recipient_name_col]).strip()
                    
                    status_text.text(f"Processing {i+1}/{total_docs}: {filename} → {rec_name}")
                    
                    if not rec_email or "@" not in rec_email:
                        results["failed"] += 1
                        results["details"].append({"File": filename, "Status": "❌ Skipped", "Error": "Invalid Email"})
                        continue
                        
                    try:
                        # Identify Attachments
                        current_batch_attachments = []
                        
                        # Find correct row index for batch matching
                        actual_row_idx = None
                        # Efficient lookup if possible, but fallback to linear search for safety
                        for r_idx in range(len(data_df)):
                             # Compare using string representation to match display logic
                             if (str(data_df.iloc[r_idx][recipient_email_col]).strip() == rec_email and 
                                 str(data_df.iloc[r_idx][recipient_name_col]).strip() == rec_name):
                                 actual_row_idx = r_idx
                                 break
                        
                        # Fallback: if data is identical order
                        if actual_row_idx is None: actual_row_idx = i

                        if "batch_attachments" in st.session_state and actual_row_idx in st.session_state.batch_attachments:
                            batch_files = st.session_state.batch_attachments[actual_row_idx]
                            # batch_files is list of (name, bytes)
                            current_batch_attachments.extend(batch_files)
                        
                        # Prepare Envelope Documents
                        # Always include generated doc
                        envelope_docs = [(filename, file_data)]
                        
                        # If using DocuSign Email, we must include all attachments in the envelope itself
                        if use_docusign_email:
                            envelope_docs.extend(common_attachments_data)
                            envelope_docs.extend(current_batch_attachments)
                        else:
                            # For Embedded, usually just the main doc is signed
                            # But user might want others? Let's just sign the main code for simplicity 
                            # and attach others to the SMTP email
                            pass

                        # 1. Send Envelope / Get Link
                        subject_formatted = ds_email_subject.replace("{Filename}", filename).replace("{Name}", rec_name)
                        body_formatted = ds_email_body.replace("{Filename}", filename).replace("{Name}", rec_name).replace("{Signing_Link}", "") # Link added later if needed checks

                        signing_url, envelope_id = ds_handler.send_envelope(
                            rec_email, 
                            rec_name, 
                            envelope_docs,
                            subject=subject_formatted,
                            body=body_formatted,
                            embedded=not use_docusign_email,
                            cc_emails=cc_list # DocuSign handles CC
                        )
                        
                        if use_docusign_email:
                            # Done! DocuSign sent the email.
                            results["sent"] += 1
                            results["details"].append({"File": filename, "Status": "✅ Sent (DocuSign)", "Envelope ID": envelope_id})
                            
                        else:
                            # Send via SMTP
                            if not signing_url:
                                raise Exception("Failed to generate signing link")
                                
                            # Prepare SMTP Attachments
                            smtp_attachments = []
                            # We don't attach the Main Doc to SMTP usually if it's being signed, 
                            # but user might want a "Review Copy". 
                            smtp_attachments.append((f"Review_{filename}", file_data))
                            smtp_attachments.extend(common_attachments_data)
                            smtp_attachments.extend(current_batch_attachments)
                            
                            # Format Link
                            signing_link_html = f'<a href="{signing_url}" style="background:#0078d4;color:white;padding:10px 20px;text-decoration:none;border-radius:4px;">Click here to sign</a>'
                            final_body_smtp = ds_email_body.replace("{Name}", rec_name).replace("{Signing_Link}", signing_link_html)
                            
                            email_handler = st.session_state.email_handler
                            success, msg = email_handler.send_personalized_email(
                                to_email=rec_email, 
                                subject=subject_formatted, 
                                body=final_body_smtp, 
                                attachment_filename=None, # Passed in additional_attachments
                                attachment_data=None,
                                cc_emails=cc_list, # Send CC via SMTP as well? Yes, but duplicates Docusign CC?
                                # If DocuSign was embedded, DocuSign did NOT send email to CC either (usually).
                                # So we handle CC here.
                                bcc_emails=bcc_list,
                                additional_attachments=smtp_attachments
                            )
                            
                            if success:
                                results["sent"] += 1
                                results["details"].append({"File": filename, "Status": "✅ Sent (SMTP)", "Envelope ID": envelope_id})
                            else:
                                results["failed"] += 1
                                results["details"].append({"File": filename, "Status": "❌ SMTP Failed", "Error": msg})

                    except Exception as e:
                        results["failed"] += 1
                        results["details"].append({"File": filename, "Status": "❌ Error", "Error": str(e)[:100]})

                    # Rate limit cushion
                    if i < total_docs - 1:
                        import time
                        time.sleep(1.0 if use_docusign_email else 0.5)
                
            st.session_state.docusign_results = results
            st.rerun()





if __name__ == "__main__":
    main()
