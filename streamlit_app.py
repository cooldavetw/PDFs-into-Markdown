import streamlit as st
from docling.document_converter import DocumentConverter
import tempfile
import os
import logging
import io
import zipfile
from pathlib import PurePosixPath

# Configure logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)


def convert_pdf_to_markdown(pdf_path):
    result = st.session_state.converter.convert(pdf_path)
    return result.document.export_to_markdown()


def convert_uploaded_pdf(uploaded_file):
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
        tmp_file.write(uploaded_file.getvalue())
        tmp_path = tmp_file.name
        logger.debug(f"Temporary file created at: {tmp_path}")

    try:
        return convert_pdf_to_markdown(tmp_path)
    finally:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)
            logger.debug("Temporary file deleted")


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

    with zipfile.ZipFile(io.BytesIO(uploaded_file.getvalue())) as input_zip:
        pdf_members = [
            info for info in input_zip.infolist()
            if not info.is_dir() and info.filename.lower().endswith('.pdf')
        ]

        if not pdf_members:
            raise ValueError("The ZIP file does not contain any PDF files.")

        with zipfile.ZipFile(markdown_zip_buffer, mode='w', compression=zipfile.ZIP_DEFLATED) as output_zip:
            for member in pdf_members:
                with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
                    tmp_file.write(input_zip.read(member))
                    tmp_path = tmp_file.name
                    logger.debug(f"Temporary ZIP PDF created at: {tmp_path}")

                try:
                    markdown_text = convert_pdf_to_markdown(tmp_path)
                    output_name = markdown_name_for_zip_member(member.filename, used_names)
                    output_zip.writestr(output_name, markdown_text)
                    converted_count += 1
                finally:
                    if os.path.exists(tmp_path):
                        os.unlink(tmp_path)
                        logger.debug("Temporary ZIP PDF deleted")

    markdown_zip_buffer.seek(0)
    return markdown_zip_buffer.getvalue(), converted_count

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

st.title("PDF to Markdown Converter")

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
    "Upload a PDF file or a ZIP file containing PDFs",
    type=['pdf', 'zip'],
    key='pdf_uploader',
    help="Drag and drop or click to select one PDF file or one ZIP file containing PDF files (max 200MB)"
)

# URL input area with spacing
st.markdown("<br>", unsafe_allow_html=True)
url = st.text_input("Or enter a PDF URL")

# Unified convert button
convert_clicked = st.button("Convert to Markdown", type="primary")

# Process either uploaded file or URL
if convert_clicked:
    if uploaded_file is not None:
        file_extension = os.path.splitext(uploaded_file.name)[1].lower()

        if file_extension == '.pdf':
            try:
                with st.spinner('Converting file...'):
                    markdown_text = convert_uploaded_pdf(uploaded_file)
                    output_filename = os.path.splitext(uploaded_file.name)[0] + '.md'

                    st.success("Conversion completed!")
                    st.download_button(
                        label="Download Markdown file",
                        data=markdown_text,
                        file_name=output_filename,
                        mime="text/markdown"
                    )

            except Exception as e:
                logger.error(f"Error processing PDF file: {str(e)}")
                st.error(f"Error processing PDF file: {str(e)}")

        elif file_extension == '.zip':
            try:
                with st.spinner('Converting PDF files from ZIP...'):
                    markdown_zip, converted_count = convert_uploaded_zip(uploaded_file)
                    output_filename = os.path.splitext(uploaded_file.name)[0] + '_markdown.zip'

                    st.success(f"Conversion completed! Converted {converted_count} PDF file(s).")
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
            st.error("Please upload a PDF file or a ZIP file containing PDFs.")
            
    elif url:
        try:
            with st.spinner('Converting from URL...'):
                logger.debug(f"Converting from URL: {url}")
                result = st.session_state.converter.convert(url)
                markdown_text = result.document.export_to_markdown()
                
                output_filename = url.split('/')[-1].split('.')[0] + '.md'
                
                st.success("Conversion completed!")
                st.download_button(
                    label="Download Markdown file",
                    data=markdown_text,
                    file_name=output_filename,
                    mime="text/markdown"
                )

        except Exception as e:
            logger.error(f"Error converting from URL: {str(e)}")
            st.error(f"Error converting from URL: {str(e)}")
    else:
        st.warning("Please upload a file or enter a URL first")
