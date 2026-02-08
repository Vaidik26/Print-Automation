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
    """Render email configuration and sending section."""
    if st.session_state.step < 4 or not st.session_state.generated_docs:
        st.info(
            "👆 Please generate documents first. Click 'Back' to go to the previous step."
        )
        render_nav_buttons(4, can_proceed=False, show_next=False)
        return

    st.markdown("### 📧 Step 5: Send Personalized Emails")

    # SMTP Configuration in sidebar
    with st.sidebar:
        st.markdown("## 📮 SMTP Configuration")
        
        smtp_server = st.text_input(
            "SMTP Server",
            value="smtp.gmail.com",
            help="Gmail: smtp.gmail.com, Outlook: smtp.office365.com",
        )
        smtp_port = st.number_input("SMTP Port", value=587, min_value=1, max_value=65535)
        sender_email = st.text_input("Sender Email", value="")
        sender_password = st.text_input(
            "App Password",
            type="password",
            help="For Gmail, generate at: https://myaccount.google.com/apppasswords",
        )
        sender_name = st.text_input("Sender Name (Optional)", value="")

        if st.button("🔌 Test Connection", use_container_width=True):
            if not sender_email or not sender_password:
                st.error("Please enter email and password")
            else:
                with st.spinner("Testing connection..."):
                    handler = EmailHandler(
                        smtp_server, smtp_port, sender_email, sender_password, sender_name
                    )
                    success, message = handler.test_connection()
                    if success:
                        st.success(message)
                        st.session_state.email_handler = handler
                        st.session_state.smtp_configured = True
                    else:
                        st.error(message)
                        st.session_state.smtp_configured = False

    # Main email section
    if not st.session_state.smtp_configured:
        st.warning("⚠️ Please configure and test SMTP connection in the sidebar first.")
        render_nav_buttons(4, can_proceed=False, show_next=False)
        return

    # Email column selection
    st.markdown("#### 1️⃣ Select Email Column")
    
    handler = st.session_state.data_handler
    columns = handler.get_columns()
    
    # Auto-detect email column
    email_column_suggestions = [col for col in columns if 'email' in col.lower() or 'e-mail' in col.lower() or 'mail' in col.lower()]
    default_email_col = email_column_suggestions[0] if email_column_suggestions else columns[0]
    
    if st.session_state.email_column is None:
        st.session_state.email_column = default_email_col
    
    email_column = st.selectbox(
        "Column containing recipient email addresses",
        options=columns,
        index=columns.index(st.session_state.email_column) if st.session_state.email_column in columns else 0,
        key="email_column_select",
    )
    st.session_state.email_column = email_column

    # Validate emails button
    if st.button("🔍 Validate Emails", type="primary"):
        data_df = handler.df
        
        validation_results = {
            "valid": [],
            "missing": [],
            "invalid": [],
        }
        
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
        
        st.session_state.email_validation_results = validation_results

    # Display validation results
    if st.session_state.email_validation_results:
        results = st.session_state.email_validation_results
        
        st.markdown("---")
        st.markdown("#### 2️⃣ Email Validation Results")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("✅ Valid", len(results["valid"]))
        with col2:
            st.metric("⚠️ Missing", len(results["missing"]))
        with col3:
            st.metric("❌ Invalid", len(results["invalid"]))

        # Missing emails section
        if results["missing"]:
            st.markdown("---")
            st.markdown("##### ⚠️ Missing Emails")
            st.markdown(
                """
                <div class="info-box">
                    <strong>Action Required:</strong> Enter email addresses for rows with missing emails, or skip them.
                </div>
                """,
                unsafe_allow_html=True,
            )
            
            # Get first few columns for display (excluding email column)
            display_columns = [col for col in columns[:3] if col != email_column]
            
            for missing_item in results["missing"][:20]:  # Show first 20
                row_idx = missing_item["row_index"]
                row_data = missing_item["row_data"]
                
                col1, col2, col3 = st.columns([2, 3, 1])
                
                with col1:
                    # Display some identifying info
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
                    skip = st.checkbox("Skip", key=f"skip_{row_idx}")
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
        
        subject_template = st.text_input(
            "Email Subject",
            value=st.session_state.email_subject_template or "Document for {" + available_placeholders[0] + "}",
            help="Use {placeholder} format",
        )
        st.session_state.email_subject_template = subject_template
        
        body_template = st.text_area(
            "Email Body",
            value=st.session_state.email_body_template or f"Dear Recipient,\n\nPlease find your document attached.\n\nBest regards",
            height=200,
            help="Use {placeholder} format",
        )
        st.session_state.email_body_template = body_template

        # Preview section
        with st.expander("👁️ Preview First Email", expanded=False):
            if results["valid"]:
                first_valid = results["valid"][0]
                first_row_idx = first_valid["row_index"]
                first_email = first_valid["email"]
                
                # Get row data
                data_df = handler.df
                first_row_data = data_df.iloc[first_row_idx].to_dict()
                
                # Render templates
                rendered_subject = EmailHandler.render_template(subject_template, first_row_data)
                rendered_body = EmailHandler.render_template(body_template, first_row_data)
                
                st.markdown(f"**To:** {first_email}")
                st.markdown(f"**Subject:** {rendered_subject}")
                st.markdown("**Body:**")
                st.text(rendered_body)
                st.markdown(f"**Attachment:** {st.session_state.generated_docs[first_row_idx][0]}")

        # Send emails section
        st.markdown("---")
        st.markdown("#### 4️⃣ Send Emails")
        
        # Calculate sendable emails
        total_valid = len(results["valid"])
        total_missing_filled = len([k for k in st.session_state.missing_emails.keys() if k not in st.session_state.skip_rows])
        total_skipped = len(st.session_state.skip_rows)
        total_sendable = total_valid + total_missing_filled
        
        st.markdown(
            f"""
            <div class="info-box">
                <strong>Ready to send:</strong> {total_sendable} emails<br>
                <strong>Skipped:</strong> {total_skipped} rows
            </div>
            """,
            unsafe_allow_html=True,
        )
        
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            if st.button("📧 Send All Emails", type="primary", use_container_width=True):
                if total_sendable == 0:
                    st.error("No valid emails to send!")
                else:
                    # Prepare email data
                    email_data_list = []
                    data_df = handler.df
                    generated_docs = st.session_state.generated_docs
                    
                    # Add valid emails
                    for valid_item in results["valid"]:
                        row_idx = valid_item["row_idx"]
                        if row_idx in st.session_state.skip_rows:
                            continue
                        
                        row_data = data_df.iloc[row_idx].to_dict()
                        email_data_list.append({
                            "to_email": valid_item["email"],
                            "subject": EmailHandler.render_template(subject_template, row_data),
                            "body": EmailHandler.render_template(body_template, row_data),
                            "attachment_filename": generated_docs[row_idx][0],
                            "attachment_data": generated_docs[row_idx][1],
                            "row_index": row_idx,
                        })
                    
                    # Add manually filled emails
                    for row_idx, manual_email in st.session_state.missing_emails.items():
                        if row_idx in st.session_state.skip_rows:
                            continue
                        if not EmailHandler.validate_email(manual_email):
                            continue
                        
                        row_data = data_df.iloc[row_idx].to_dict()
                        email_data_list.append({
                            "to_email": manual_email,
                            "subject": EmailHandler.render_template(subject_template, row_data),
                            "body": EmailHandler.render_template(body_template, row_data),
                            "attachment_filename": generated_docs[row_idx][0],
                            "attachment_data": generated_docs[row_idx][1],
                            "row_index": row_idx,
                        })
                    
                    # Send emails with progress
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    def progress_callback(current, total, status):
                        progress_bar.progress(current / total)
                        status_text.text(f"{status} ({current}/{total})")
                    
                    email_handler = st.session_state.email_handler
                    send_results = email_handler.send_batch_emails(
                        email_data_list, progress_callback=progress_callback
                    )
                    
                    progress_bar.empty()
                    status_text.empty()
                    
                    # Show results
                    st.markdown("---")
                    st.markdown("### ✅ Email Sending Complete!")
                    
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("✅ Sent", send_results["sent"])
                    with col2:
                        st.metric("❌ Failed", send_results["failed"])
                    with col3:
                        st.metric("⏭️ Skipped", total_skipped)
                    
                    if send_results["failed"] > 0:
                        st.markdown("##### ❌ Failed Emails")
                        for failed in send_results["failed_details"]:
                            st.error(f"Row {failed['row_index'] + 1} ({failed['email']}): {failed['error']}")
                    
                    st.session_state.step = max(st.session_state.step, 6)

    # Navigation buttons
    render_nav_buttons(4, can_proceed=False, show_next=False)



def render_email_section():
    """Render email configuration and sending section."""
    if st.session_state.step < 4 or not st.session_state.generated_docs:
        st.info(
            "👆 Please generate documents first. Click 'Back' to go to the previous step."
        )
        render_nav_buttons(4, can_proceed=False, show_next=False)
        return

    st.markdown("### 📧 Step 5: Send Personalized Emails")

    # SMTP Configuration on main page
    st.markdown("#### 1️⃣ SMTP Configuration")
    
    st.markdown(
        """
        <div class="info-box">
            <strong>📮 Configure your email settings to send documents</strong>
        </div>
        """,
        unsafe_allow_html=True,
    )
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        # Email input with provider detection
        sender_email = st.text_input(
            "Your Email Address",
            value="",
            placeholder="your.email@gmail.com",
            help="Enter your email address to send from"
        )
    
    with col2:
        sender_name = st.text_input(
            "Sender Name (Optional)",
            value="",
            placeholder="Your Name or Company Name"
        )
    
    # Auto-detect provider and configure SMTP
    smtp_server = "smtp.gmail.com"
    smtp_port = 587
    
    if sender_email:
        if "gmail.com" in sender_email.lower():
            smtp_server = "smtp.gmail.com"
            smtp_port = 587
            provider = "Gmail"
        elif "outlook.com" in sender_email.lower() or "hotmail.com" in sender_email.lower():
            smtp_server = "smtp.office365.com"
            smtp_port = 587
            provider = "Outlook"
        elif "yahoo.com" in sender_email.lower():
            smtp_server = "smtp.mail.yahoo.com"
            smtp_port = 587
            provider = "Yahoo"
        else:
            provider = "Custom"
            col1, col2 = st.columns(2)
            with col1:
                smtp_server = st.text_input("SMTP Server", value="smtp.gmail.com")
            with col2:
                smtp_port = st.number_input("SMTP Port", value=587, min_value=1, max_value=65535)
        
        if provider in ["Gmail", "Outlook", "Yahoo"]:
            st.success(f"✅ Detected: {provider} (SMTP: {smtp_server}:{smtp_port})")
    
    # App Password with direct link
    if sender_email and "gmail.com" in sender_email.lower():
        col1, col2 = st.columns([3, 2])
        with col1:
            st.markdown(
                """
                <div style="background: rgba(102, 126, 234, 0.1); padding: 10px; border-radius: 8px; margin-bottom: 10px;">
                    <strong>📝 How to get Gmail App Password:</strong><br>
                    1. Click the link → 
                    2. Sign in to your Google Account<br>
                    3. Create an App Password<br>
                    4. Copy and paste it below
                </div>
                """,
                unsafe_allow_html=True
            )
        with col2:
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown(
                "🔗 [**Get Gmail App Password →**](https://myaccount.google.com/apppasswords)",
                unsafe_allow_html=True
            )
    elif sender_email and "outlook.com" in sender_email.lower():
        st.info("💡 For Outlook, you may need to enable 'Less secure app access' or use an App Password")
    elif sender_email and "yahoo.com" in sender_email.lower():
        st.info("💡 For Yahoo, generate an App Password from Account Security settings")
    
    col1, col2 = st.columns([3, 1])
    with col1:
        sender_password = st.text_input(
            "App Password",
            type="password",
            placeholder="Enter your app password here",
            help="Paste the app password you generated"
        )
    with col2:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🔌 Test Connection", use_container_width=True):
            if not sender_email or not sender_password:
                st.error("❌ Please enter both email and app password")
            else:
                with st.spinner("Testing connection..."):
                    handler = EmailHandler(
                        smtp_server, smtp_port, sender_email, sender_password, sender_name
                    )
                    success, message = handler.test_connection()
                    if success:
                        st.success(message)
                        st.session_state.email_handler = handler
                        st.session_state.smtp_configured = True
                    else:
                        st.error(message)
                        st.session_state.smtp_configured = False

    # Main email section
    if not st.session_state.smtp_configured:
        st.warning("⚠️ Please configure and test SMTP connection in the sidebar first.")
        render_nav_buttons(4, can_proceed=False, show_next=False)
        return

    # Email column selection
    st.markdown("---")
    st.markdown("#### 2️⃣ Select Email Column")
    
    handler = st.session_state.data_handler
    columns = handler.get_columns()
    
    # Auto-detect email column
    email_column_suggestions = [col for col in columns if 'email' in col.lower() or 'e-mail' in col.lower() or 'mail' in col.lower()]
    default_email_col = email_column_suggestions[0] if email_column_suggestions else columns[0]
    
    if st.session_state.email_column is None:
        st.session_state.email_column = default_email_col
    
    email_column = st.selectbox(
        "Column containing recipient email addresses",
        options=columns,
        index=columns.index(st.session_state.email_column) if st.session_state.email_column in columns else 0,
        key="email_column_select",
    )
    st.session_state.email_column = email_column

    # Validate emails button
    if st.button("🔍 Validate Emails", type="primary"):
        data_df = handler.df
        
        validation_results = {
            "valid": [],
            "missing": [],
            "invalid": [],
        }
        
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
        
        st.session_state.email_validation_results = validation_results

    # Display validation results
    if st.session_state.email_validation_results:
        results = st.session_state.email_validation_results
        
        st.markdown("---")
        st.markdown("#### 3️⃣ Email Validation Results")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("✅ Valid", len(results["valid"]))
        with col2:
            st.metric("⚠️ Missing", len(results["missing"]))
        with col3:
            st.metric("❌ Invalid", len(results["invalid"]))

        # Missing emails section
        if results["missing"]:
            st.markdown("---")
            st.markdown("##### ⚠️ Missing Emails")
            st.markdown(
                """
                <div class="info-box">
                    <strong>Action Required:</strong> Enter email addresses for rows with missing emails, or skip them.
                </div>
                """,
                unsafe_allow_html=True,
            )
            
            # Bulk action buttons
            col1, col2, col3 = st.columns([1, 1, 3])
            with col1:
                if st.button("✅ Select All", use_container_width=True, help="Select all rows with missing emails"):
                    # Clear skip_rows for all missing email rows
                    for missing_item in results["missing"]:
                        row_idx = missing_item["row_index"]
                        if row_idx in st.session_state.skip_rows:
                            st.session_state.skip_rows.remove(row_idx)
                    st.success(f"Selected all {len(results['missing'])} rows")
                    st.rerun()
            
            with col2:
                if st.button("⏭️ Skip All", use_container_width=True, help="Skip all rows with missing emails"):
                    # Add all missing email rows to skip_rows
                    for missing_item in results["missing"]:
                        row_idx = missing_item["row_index"]
                        st.session_state.skip_rows.add(row_idx)
                    st.success(f"Skipped all {len(results['missing'])} rows")
                    st.rerun()
            
            # Get first few columns for display (excluding email column)
            display_columns = [col for col in columns[:3] if col != email_column]
            
            for missing_item in results["missing"][:20]:  # Show first 20
                row_idx = missing_item["row_index"]
                row_data = missing_item["row_data"]
                
                col1, col2, col3 = st.columns([2, 3, 1])
                
                with col1:
                    # Display some identifying info
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
                value=st.session_state.cc_emails,
                placeholder="email1@example.com, email2@example.com",
                help="Separate multiple emails with commas"
            )
            st.session_state.cc_emails = cc_emails
        
        with col2:
            bcc_emails = st.text_input(
                "BCC (Optional)",
                value=st.session_state.bcc_emails,
                placeholder="email1@example.com, email2@example.com",
                help="Separate multiple emails with commas"
            )
            st.session_state.bcc_emails = bcc_emails
        
        
        subject_template = st.text_input(
            "Email Subject",
            value=st.session_state.email_subject_template or "Document for {" + available_placeholders[0] + "}",
            help="Use {placeholder} format",
        )
        st.session_state.email_subject_template = subject_template
        
        body_template = st.text_area(
            "Email Body",
            value=st.session_state.email_body_template or f"Dear Recipient,\n\nPlease find your document attached.\n\nBest regards",
            height=200,
            help="Use {placeholder} format",
        )
        st.session_state.email_body_template = body_template
        
        # Common attachment section
        st.markdown("**📎 Common Attachment (Optional)**")
        st.markdown(
            """
            <div class="info-box">
                Upload a file that will be attached to <strong>all emails</strong> (e.g., brochure, terms & conditions, etc.)
            </div>
            """,
            unsafe_allow_html=True,
        )
        
        col1, col2 = st.columns([3, 1])
        with col1:
            common_file = st.file_uploader(
                "Upload Common Attachment",
                type=["pdf", "doc", "docx", "xls", "xlsx", "txt", "jpg", "png"],
                help="This file will be sent with every email",
                label_visibility="collapsed",
                key="common_file_uploader"
            )
            
            if common_file:
                st.session_state.common_attachment = (common_file.name, common_file.read())
                st.success(f"✅ Attached: {common_file.name} ({len(st.session_state.common_attachment[1]) / 1024:.1f} KB)")
        
        with col2:
            if st.session_state.common_attachment:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("🗑️ Remove", use_container_width=True):
                    st.session_state.common_attachment = None
                    st.rerun()

        # Preview section
        with st.expander("👁️ Preview First Email", expanded=False):
            if results["valid"]:
                first_valid = results["valid"][0]
                first_row_idx = first_valid["row_index"]
                first_email = first_valid["email"]
                
                # Get row data
                data_df = handler.df
                first_row_data = data_df.iloc[first_row_idx].to_dict()
                
                # Render templates
                rendered_subject = EmailHandler.render_template(subject_template, first_row_data)
                rendered_body = EmailHandler.render_template(body_template, first_row_data)
                
                st.markdown(f"**To:** {first_email}")
                st.markdown(f"**Subject:** {rendered_subject}")
                st.markdown("**Body:**")
                st.text(rendered_body)
                st.markdown(f"**Attachment:** {st.session_state.generated_docs[first_row_idx][0]}")

        # Preview Recipients section
        st.markdown("---")
        st.markdown("#### 5️⃣ Preview Recipients")
        
        st.markdown(
            """
            <div class="info-box">
                <strong>📋 Review who will receive emails and what will be sent</strong>
            </div>
            """,
            unsafe_allow_html=True,
        )
        
        # Calculate sendable emails
        total_valid = len(results["valid"])
        total_missing_filled = len([k for k in st.session_state.missing_emails.keys() if k not in st.session_state.skip_rows])
        total_skipped = len(st.session_state.skip_rows)
        total_sendable = total_valid + total_missing_filled
        
        # Summary metrics
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("📧 Total Recipients", total_sendable)
        with col2:
            st.metric("⏭️ Skipped", total_skipped)
        with col3:
            if st.session_state.cc_emails:
                cc_count = len([e for e in st.session_state.cc_emails.split(',') if e.strip()])
                st.metric("📎 CC", cc_count)
            else:
                st.metric("📎 CC", 0)
        with col4:
            if st.session_state.common_attachment:
                st.metric("📄 Common File", "Yes")
            else:
                st.metric("📄 Common File", "No")
        
        if total_sendable == 0:
            st.warning("⚠️ No recipients to send emails to. Please add valid emails or uncheck skip boxes.")
        else:
            # Build preview data
            preview_data = []
            data_df = handler.df
            generated_docs = st.session_state.generated_docs
            
            # Add valid emails
            for valid_item in results["valid"]:
                row_idx = valid_item["row_index"]
                if row_idx in st.session_state.skip_rows:
                    continue
                
                row_data = data_df.iloc[row_idx].to_dict()
                # Get first few columns for display
                display_cols = list(row_data.keys())[:3]
                display_info = " | ".join([f"{col}: {row_data.get(col, 'N/A')}" for col in display_cols])
                
                preview_data.append({
                    "Row": row_idx + 1,
                    "Recipient Info": display_info,
                    "Email": valid_item["email"],
                    "Document": generated_docs[row_idx][0],
                })
            
            # Add manually filled emails
            for row_idx, manual_email in st.session_state.missing_emails.items():
                if row_idx in st.session_state.skip_rows:
                    continue
                if not EmailHandler.validate_email(manual_email):
                    continue
                
                row_data = data_df.iloc[row_idx].to_dict()
                display_cols = list(row_data.keys())[:3]
                display_info = " | ".join([f"{col}: {row_data.get(col, 'N/A')}" for col in display_cols])
                
                preview_data.append({
                    "Row": row_idx + 1,
                    "Recipient Info": display_info,
                    "Email": manual_email,
                    "Document": generated_docs[row_idx][0],
                })
            
            # Display preview table
            if preview_data:
                import pandas as pd
                preview_df = pd.DataFrame(preview_data)
                st.markdown("**📋 Recipients List:**")
                st.dataframe(preview_df, use_container_width=True, hide_index=True)
                
                # Show CC/BCC if configured
                if st.session_state.cc_emails or st.session_state.bcc_emails:
                    st.markdown("**Additional Recipients:**")
                    if st.session_state.cc_emails:
                        st.markdown(f"- **CC:** {st.session_state.cc_emails}")
                    if st.session_state.bcc_emails:
                        st.markdown(f"- **BCC:** {st.session_state.bcc_emails}")
                
                # Show common attachment if uploaded
                if st.session_state.common_attachment:
                    st.markdown(f"**📎 Common Attachment:** {st.session_state.common_attachment[0]} ({len(st.session_state.common_attachment[1]) / 1024:.1f} KB)")
                
                # Confirm and send button
                st.markdown("---")
                
                # Speed control
                st.markdown("**⏱️ Sending Speed:**")
                col_a, col_b = st.columns([3, 1])
                with col_a:
                    delay_seconds = st.slider(
                        "Delay between emails (seconds)",
                        min_value=0.1,
                        max_value=2.0,
                        value=0.5,
                        step=0.1,
                        help="Lower = faster, but may hit rate limits. Gmail recommended: 0.5-1.0s",
                        label_visibility="collapsed"
                    )
                with col_b:
                    estimated_time = total_sendable * delay_seconds
                    st.metric("Est. Time", f"{estimated_time:.0f}s")
                
                st.info(f"💡 **Tip:** {delay_seconds}s delay = ~{60/delay_seconds:.0f} emails/minute. Gmail limit: ~500 emails/day.")
                
                col1, col2, col3 = st.columns([1, 2, 1])
                with col2:
                    if st.button("✅ Confirm & Send All Emails", type="primary", use_container_width=True):
                        # Prepare email data
                        email_data_list = []
                        
                        # Add valid emails
                        for valid_item in results["valid"]:
                            row_idx = valid_item["row_index"]
                            if row_idx in st.session_state.skip_rows:
                                continue
                            
                            row_data = data_df.iloc[row_idx].to_dict()
                            email_data_list.append({
                                "to_email": valid_item["email"],
                                "subject": EmailHandler.render_template(subject_template, row_data),
                                "body": EmailHandler.render_template(body_template, row_data),
                                "attachment_filename": generated_docs[row_idx][0],
                                "attachment_data": generated_docs[row_idx][1],
                                "cc_emails": [e.strip() for e in st.session_state.cc_emails.split(',') if e.strip()] if st.session_state.cc_emails else None,
                                "bcc_emails": [e.strip() for e in st.session_state.bcc_emails.split(',') if e.strip()] if st.session_state.bcc_emails else None,
                                "additional_attachments": [st.session_state.common_attachment] if st.session_state.common_attachment else None,
                                "row_index": row_idx,
                                "recipient_info": " | ".join([f"{col}: {row_data.get(col, 'N/A')}" for col in list(row_data.keys())[:3]]),
                            })
                        
                        # Add manually filled emails
                        for row_idx, manual_email in st.session_state.missing_emails.items():
                            if row_idx in st.session_state.skip_rows:
                                continue
                            if not EmailHandler.validate_email(manual_email):
                                continue
                            
                            row_data = data_df.iloc[row_idx].to_dict()
                            email_data_list.append({
                                "to_email": manual_email,
                                "subject": EmailHandler.render_template(subject_template, row_data),
                                "body": EmailHandler.render_template(body_template, row_data),
                                "attachment_filename": generated_docs[row_idx][0],
                                "attachment_data": generated_docs[row_idx][1],
                                "cc_emails": [e.strip() for e in st.session_state.cc_emails.split(',') if e.strip()] if st.session_state.cc_emails else None,
                                "bcc_emails": [e.strip() for e in st.session_state.bcc_emails.split(',') if e.strip()] if st.session_state.bcc_emails else None,
                                "additional_attachments": [st.session_state.common_attachment] if st.session_state.common_attachment else None,
                                "row_index": row_idx,
                                "recipient_info": " | ".join([f"{col}: {row_data.get(col, 'N/A')}" for col in list(row_data.keys())[:3]]),
                            })
                        
                        # Send emails with progress
                        progress_bar = st.progress(0)
                        status_text = st.empty()
                        
                        def progress_callback(current, total, status):
                            progress_bar.progress(current / total)
                            status_text.text(f"{status} ({current}/{total})")
                        
                        email_handler = st.session_state.email_handler
                        send_results = email_handler.send_batch_emails(
                            email_data_list,
                            progress_callback=progress_callback,
                            delay_seconds=delay_seconds  # Use user-selected delay
                        )
                        
                        progress_bar.empty()
                        status_text.empty()
                        
                        # Store results with detailed info
                        detailed_results = []
                        for email_data in email_data_list:
                            row_idx = email_data["row_index"]
                            # Check if this email was in failed list
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
                        
                        st.session_state.step = max(st.session_state.step, 6)
                        st.rerun()
        
        # Show results if available
        if st.session_state.email_send_results:
            st.markdown("---")
            st.markdown("#### 6️⃣ Send Results")
            
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
            
            # Download results as CSV
            csv = results_df.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Download Results as CSV",
                data=csv,
                file_name="email_send_results.csv",
                mime="text/csv",
            )
            
            # Reset button
            if st.button("🔄 Send More Emails"):
                st.session_state.email_send_results = None
                st.rerun()


    # Navigation buttons
    render_nav_buttons(4, can_proceed=False, show_next=False)



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
    tab_names = ["📝 Template", "📊 Data", "🔗 Mapping", "🚀 Generate", "📧 Email"]

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


if __name__ == "__main__":
    main()
