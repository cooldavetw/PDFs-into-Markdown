import streamlit as st
from docling.document_converter import DocumentConverter
import tempfile
import os
import logging
import io
import zipfile
import shutil
import subprocess
from pathlib import Path, PurePosixPath

# Configure logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

DIRECT_CONVERSION_EXTENSIONS = ('.pdf', '.docx')
SUPPORTED_DOCUMENT_EXTENSIONS = DIRECT_CONVERSION_EXTENSIONS + ('.doc',)


def convert_document_to_markdown(document_path):
    result = st.session_state.converter.convert(document_path)
    return result.document.export_to_markdown()


def convert_doc_to_docx(doc_path):
    office_command = shutil.which('soffice') or shutil.which('libreoffice')
    if office_command is None:
        raise RuntimeError("LibreOffice is required to convert legacy .doc files.")

    with tempfile.TemporaryDirectory() as output_dir, tempfile.TemporaryDirectory() as profile_dir:
        libreoffice_env = os.environ.copy()
        libreoffice_env["HOME"] = profile_dir

        completed_process = subprocess.run(
            [
                office_command,
                f"-env:UserInstallation={Path(profile_dir).as_uri()}",
                '--headless',
                '--convert-to',
                'docx',
                '--outdir',
                output_dir,
                doc_path,
            ],
            capture_output=True,
            text=True,
            timeout=120,
            check=False,
            env=libreoffice_env,
        )

        output_path = os.path.join(
            output_dir,
            os.path.splitext(os.path.basename(doc_path))[0] + '.docx',
        )
        if completed_process.returncode != 0 or not os.path.exists(output_path):
            error_output = completed_process.stderr or completed_process.stdout
            raise RuntimeError(f"Error converting DOC file: {error_output.strip()}")

        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
            with open(output_path, 'rb') as converted_file:
                shutil.copyfileobj(converted_file, tmp_file)
            return tmp_file.name


def convert_uploaded_document(uploaded_file):
    file_extension = os.path.splitext(uploaded_file.name)[1].lower()

    with tempfile.NamedTemporaryFile(delete=False, suffix=file_extension) as tmp_file:
        tmp_file.write(uploaded_file.getvalue())
        tmp_path = tmp_file.name
        logger.debug(f"Temporary file created at: {tmp_path}")

    converted_docx_path = None
    try:
        conversion_path = tmp_path
        if file_extension == '.doc':
            converted_docx_path = convert_doc_to_docx(tmp_path)
            conversion_path = converted_docx_path

        return convert_document_to_markdown(conversion_path)
    finally:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)
            logger.debug("Temporary file deleted")
        if converted_docx_path and os.path.exists(converted_docx_path):
            os.unlink(converted_docx_path)
            logger.debug("Temporary converted DOCX file deleted")


def is_supported_document(filename):
    return filename.lower().endswith(SUPPORTED_DOCUMENT_EXTENSIONS)


def markdown_name_for_zip_member(member_name, used_names):
    path = PurePosixPath(member_name.replace('\\', '/'))
    safe_parts = [
        part for part in path.parts
        if part not in ('', '.', '..') and not part.startswith('/')
    ]
    path = PurePosixPath(*safe_parts) if safe_parts else PurePosixPath('document.pdf')
    output_name = str(path.with_suffix('.md'))

    if output_name not in used_names:
        used_names.add(output_name)
        return output_name

    stem = str(path.with_suffix(''))
    counter = 2
    while True:
        candidate = f"{stem}-{counter}.md"
        if candidate not in used_names:
            used_names.add(candidate)
            return candidate
        counter += 1


def convert_uploaded_zip(uploaded_file):
    markdown_zip_buffer = io.BytesIO()
    converted_count = 0
    used_names = set()
    markdown_files = []
    uploaded_document_names = []

    with zipfile.ZipFile(io.BytesIO(uploaded_file.getvalue())) as input_zip:
        document_members = [
            info for info in input_zip.infolist()
            if not info.is_dir() and is_supported_document(info.filename)
        ]

        if not document_members:
            raise ValueError("The ZIP file does not contain any supported document files.")

        with zipfile.ZipFile(markdown_zip_buffer, mode='w', compression=zipfile.ZIP_DEFLATED) as output_zip:
            for member in document_members:
                file_extension = os.path.splitext(member.filename)[1].lower()
                uploaded_document_names.append(member.filename)

                with tempfile.NamedTemporaryFile(delete=False, suffix=file_extension) as tmp_file:
                    tmp_file.write(input_zip.read(member))
                    tmp_path = tmp_file.name
                    logger.debug(f"Temporary ZIP document created at: {tmp_path}")

                converted_docx_path = None
                try:
                    conversion_path = tmp_path
                    if file_extension == '.doc':
                        converted_docx_path = convert_doc_to_docx(tmp_path)
                        conversion_path = converted_docx_path

                    markdown_text = convert_document_to_markdown(conversion_path)
                    output_name = markdown_name_for_zip_member(member.filename, used_names)
                    output_zip.writestr(output_name, markdown_text)
                    markdown_files.append((output_name, markdown_text))
                    converted_count += 1
                finally:
                    if os.path.exists(tmp_path):
                        os.unlink(tmp_path)
                        logger.debug("Temporary ZIP document deleted")
                    if converted_docx_path and os.path.exists(converted_docx_path):
                        os.unlink(converted_docx_path)
                        logger.debug("Temporary ZIP converted DOCX file deleted")

    markdown_zip_buffer.seek(0)
    return markdown_zip_buffer.getvalue(), converted_count, markdown_files, uploaded_document_names


def display_uploaded_documents(document_names, key):
    st.subheader("Uploaded Documents")
    st.text_area(
        "Uploaded Documents",
        "\n".join(document_names),
        height=min(300, 70 + (len(document_names) * 24)),
        key=key,
        label_visibility="collapsed",
    )


def display_markdown_preview(markdown_text, key):
    st.subheader("Converted Markdown")
    st.text_area(
        "Converted Markdown",
        markdown_text,
        height=500,
        key=key,
        label_visibility="collapsed",
    )


def display_zip_markdown_previews(markdown_files):
    st.subheader("Converted Markdown Files")
    for index, (filename, markdown_text) in enumerate(markdown_files):
        with st.expander(filename):
            st.text_area(
                "Converted Markdown",
                markdown_text,
                height=400,
                key=f"zip_markdown_preview_{index}",
                label_visibility="collapsed",
            )

# Custom CSS for better layout
st.markdown("""
    <style>    
        .stFileUploader {
            padding: 1rem;
        }
        
        button[data-testid="stFileUploaderButtonPrimary"] {
            background-color: #000660 !important;
            border: none !important;
            color: white !important;
        }

        .stButton button {
            background-color: #006666;
            border: none !important;
            color: white;
            padding: 0.5rem 2rem !important;
        }
        .stButton button:hover {
            background-color: #008080 !important;
            color: white !important;
            border-color: #008080 !important;
        }
        .upload-text {
            font-size: 1.2rem;
            margin-bottom: 1rem;
        }
        div[data-testid="stFileUploadDropzone"]:hover {
            border-color: #006666 !important;
            background-color: rgba(0, 102, 102, 0.05) !important;
        }
    </style>
""", unsafe_allow_html=True)

st.title("Document to Markdown Converter")

# Initialize session state if it doesn't exist
if 'converter' not in st.session_state:
    try:
        st.session_state.converter = DocumentConverter()
        logger.debug("Converter successfully created")
    except Exception as e:
        logger.error(f"Error creating converter: {str(e)}")
        st.error(f"Error creating converter: {str(e)}")
        st.stop()

# Main upload area
uploaded_file = st.file_uploader(
    "Upload a PDF, Word document, or ZIP file containing documents",
    type=['pdf', 'doc', 'docx', 'zip'],
    key='pdf_uploader',
    help="Drag and drop or click to select one PDF/Word file or one ZIP file containing documents (max 200MB)"
)

# Unified convert button
convert_clicked = st.button("Convert to Markdown", type="primary")

# Process uploaded file
if convert_clicked:
    if uploaded_file is not None:
        file_extension = os.path.splitext(uploaded_file.name)[1].lower()

        if file_extension in SUPPORTED_DOCUMENT_EXTENSIONS:
            try:
                with st.spinner('Converting file...'):
                    markdown_text = convert_uploaded_document(uploaded_file)
                    output_filename = os.path.splitext(uploaded_file.name)[0] + '.md'

                    st.success("Conversion completed!")
                    display_uploaded_documents([uploaded_file.name], "uploaded_document_list")
                    display_markdown_preview(markdown_text, "uploaded_document_markdown_preview")
                    st.download_button(
                        label="Download Markdown file",
                        data=markdown_text,
                        file_name=output_filename,
                        mime="text/markdown"
                    )

            except Exception as e:
                logger.error(f"Error processing document file: {str(e)}")
                st.error(f"Error processing document file: {str(e)}")

        elif file_extension == '.zip':
            try:
                with st.spinner('Converting document files from ZIP...'):
                    markdown_zip, converted_count, markdown_files, uploaded_document_names = convert_uploaded_zip(uploaded_file)
                    output_filename = os.path.splitext(uploaded_file.name)[0] + '_markdown.zip'

                    st.success(f"Conversion completed! Converted {converted_count} document file(s).")
                    display_uploaded_documents(uploaded_document_names, "zip_uploaded_document_list")
                    display_zip_markdown_previews(markdown_files)
                    st.download_button(
                        label="Download Markdown ZIP file",
                        data=markdown_zip,
                        file_name=output_filename,
                        mime="application/zip"
                    )

            except zipfile.BadZipFile:
                logger.error("Uploaded file is not a valid ZIP file")
                st.error("Uploaded file is not a valid ZIP file.")
            except Exception as e:
                logger.error(f"Error processing ZIP file: {str(e)}")
                st.error(f"Error processing ZIP file: {str(e)}")

        else:
            st.error("Please upload a PDF, Word document, or ZIP file containing supported documents.")
    else:
        st.warning("Please upload a PDF, Word document, or ZIP file containing supported documents first")
