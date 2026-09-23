import io
import json
import os
import re
import tempfile
import zipfile
from PIL import Image
import docx
from docx.shared import Inches, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
import pymupdf
import pypandoc
import streamlit as st
from google import genai
from google.genai import types

# Setup Page Configuration
st.set_page_config(page_title="Document to Word Converter with Image Extraction", page_icon="📝", layout="centered")



# --- Sidebar Disclaimer & Signature Section ---
with st.sidebar:
    st.header("⚠️ Legal & Technical Disclaimer")
    st.markdown(
        """
        **Application Origins:**  
        This application was **developed in collaboration with Gemini AI** as an automated utility solution.
        
        **Experimental Purpose:**  
        This application is developed strictly for experimental, proof-of-concept, and internal evaluation purposes. 
        
        **Accuracy & Liability:**  
        * Text generation, OCR, and document extraction are processed using artificial intelligence (Gemini API).  
        * Outputs may contain **errors, hallucinations, omissions, or formatting discrepancies**.  
        * The user acknowledges and accepts all operational risks associated with using this software.  
        * The developer offers **no warranties of any kind** (express or implied) and shall **not be held liable** for any direct, indirect, incidental, or consequential damages, data loss, or business interruptions arising out of the use or inability to use this tool.
        
        **Data Privacy Notice:**  
        Uploaded files are processed entirely in temporary memory buffers and are automatically destroyed when your session ends or closes. No data is permanently retained on this server.
        """
    )
    
    st.write("---")
    st.subheader("⚙️ Settings")
    model_choice = st.selectbox(
        "Gemini Model:",
        options=["gemini-3.8-flash", "gemini-3.6-flash"],
        index=0,
        help="gemini-3.8-flash provides deep spatial reasoning and high accuracy; gemini-3.6-flash is lightweight and fast."
    )
    
    render_dpi = st.select_slider(
        "PDF Extraction Quality (DPI):",
        options=[150, 200, 300],
        value=200,
        help="Higher DPI produces crisper extracted images from PDF pages."
    )

    st.write("---")
    st.subheader("✍️ Digital Acknowledgement")
    # Interactive signature text entry
    user_signature = st.text_input(
        "Type your full name to accept these terms:", 
        placeholder="First and Last Name",
        help="Entering your name acts as a digital signature acknowledging the risks listed above."
    ).strip()

    # Track if the disclaimer has been signed
    is_signed = len(user_signature) > 0


st.markdown("<h1 style='font-size: 32px;'>📝 PNG/PDF→Markdown→Word Converter</h1>", unsafe_allow_html=True)
st.write(
    "Upload your PNG images or PDF documents. Gemini AI will extract all text, tables, equations, "
    "and **visual figures/diagrams**, embed the images directly into Markdown, and compile a fully formatted **Word Document (.docx)**. "
    "Developed by Dr. Vinu Unnikrishnan in collaboration with **Gemini AI**."
)

# Check signature status before displaying instructions or core application tools
if not is_signed:
    st.warning("🔒 Please read and sign the **Disclaimer** in the sidebar to unlock the application.")
else:
    st.success(f"✍️ Acknowledged by: **{user_signature}**")

    # --- User Instructions & API Key Guide ---
    st.markdown("### 🔑 Getting Started: How to Get Your Gemini API Key")
    st.markdown(
        """
        To use this tool, you need a **Google Gemini API key**. Getting one takes less than a minute and is completely **free** for experimental tiers:
        1. **Go to Google AI Studio:** Click on **[Google AI Studio](https://aistudio.google.com)** and log in using your Google or Workspace account.
        2. **Create Key:** Click the blue **"Get API key"** button in the dashboard (🔑).
        3. **Copy Key:** Choose **"Create API key"**, select or create a project, and copy your generated key string.
        4. **Paste Below:** Paste that key string into the field below to unlock the file processor.
        """
    )
    st.write("---")

    # --- API Key Input Field ---
    api_key_input = st.text_input(
        "Enter your Gemini API Key:", 
        type="password", 
        placeholder="Paste your Gemini API key here...",
        help="Paste the key you generated from https://aistudio.google.com"
    )
    final_key = api_key_input.strip()

    # --- File Upload Section (Accepts PNG and PDF) ---
    uploaded_files = st.file_uploader(
        "Drag and drop or browse PNG/PDF files:", 
        type=["png", "pdf", "jpg", "jpeg"], 
        accept_multiple_files=True,
        help="Hold Ctrl/Cmd to select multiple files"
    )

    # Sort files alphabetically by name to preserve document flow order
    if uploaded_files:
        uploaded_files = sorted(uploaded_files, key=lambda x: x.name)
        st.info(f"📁 Loaded {len(uploaded_files)} file(s) for conversion.")

    # --- Filename Customization ---
    default_name = "converted_document"
    output_name = st.text_input("Output Document Name (without extension):", value=default_name)

    # --- Helper Functions ---
    def prepare_file_manifest(uploaded_files):
        """Inspects uploaded files and returns manifest and Gemini contents parts."""
        files_meta = []
        gemini_parts = []
        manifest_lines = []

        for idx, uploaded_file in enumerate(uploaded_files):
            file_bytes = uploaded_file.read()
            # Reset pointer after reading
            uploaded_file.seek(0)
            fname = uploaded_file.name
            lower_name = fname.lower()

            if lower_name.endswith('.pdf'):
                mime_type = 'application/pdf'
                try:
                    doc = pymupdf.open(stream=file_bytes, filetype="pdf")
                    page_count = len(doc)
                except Exception:
                    page_count = 1
                files_meta.append({
                    "index": idx,
                    "name": fname,
                    "type": "pdf",
                    "bytes": file_bytes,
                    "page_count": page_count
                })
                manifest_lines.append(f"- File [{idx}]: \"{fname}\" (PDF document, {page_count} page(s))")
            else:
                mime_type = 'image/png' if lower_name.endswith('.png') else 'image/jpeg'
                try:
                    pil_img = Image.open(io.BytesIO(file_bytes))
                    width, height = pil_img.size
                except Exception:
                    width, height = 0, 0
                files_meta.append({
                    "index": idx,
                    "name": fname,
                    "type": "image",
                    "bytes": file_bytes,
                    "page_count": 1,
                    "size": (width, height)
                })
                manifest_lines.append(f"- File [{idx}]: \"{fname}\" (Image, dimensions {width}x{height})")

            gemini_parts.append(
                types.Part.from_bytes(data=file_bytes, mime_type=mime_type)
            )

        return files_meta, gemini_parts, "\n".join(manifest_lines)

    def parse_gemini_document_response(resp_text):
        r"""
        Robustly parses Gemini response containing figures JSON and markdown.
        Handles unescaped LaTeX backslashes without crashing on Invalid \escape.
        """
        resp_text = resp_text.strip()
        if resp_text.startswith("```json"):
            resp_text = resp_text[7:]
        elif resp_text.startswith("```"):
            resp_text = resp_text[3:]
        if resp_text.endswith("```"):
            resp_text = resp_text[:-3]
        resp_text = resp_text.strip()

        figures = []
        markdown_text = ""

        # Strategy A: Try direct json.loads
        try:
            data = json.loads(resp_text)
            figures = data.get("figures", [])
            markdown_text = data.get("markdown", "")
            if markdown_text:
                return figures, markdown_text
        except Exception:
            pass

        # Strategy B: Robust extraction of "figures" and "markdown"
        fig_match = re.search(r'"figures"\s*:\s*(\[[\s\S]*?\])\s*,\s*"markdown"', resp_text)
        if not fig_match:
            fig_match = re.search(r'"figures"\s*:\s*(\[[\s\S]*?\])', resp_text)

        if fig_match:
            fig_str = fig_match.group(1)
            # Escape invalid backslashes inside figures block
            fig_str_clean = re.sub(r'\\(?!["\\/bfnrt]|u[0-9a-fA-F]{4})', r'\\\\', fig_str)
            try:
                figures = json.loads(fig_str_clean)
            except Exception:
                obj_pattern = r'\{\s*"id"\s*:\s*"([^"]+)"[\s\S]*?"box_2d"\s*:\s*\[([^\]]+)\][\s\S]*?\}'
                for m in re.finditer(obj_pattern, fig_str):
                    try:
                        fid = m.group(1)
                        box = [int(x.strip()) for x in m.group(2).split(',')]
                        figures.append({"id": fid, "box_2d": box, "file_index": 0, "page_index": 0, "caption": fid})
                    except Exception:
                        pass

        md_match = re.search(r'"markdown"\s*:\s*"(.*)"\s*\}?\s*$', resp_text, re.DOTALL)
        if md_match:
            raw_md = md_match.group(1)
            # Decode JSON whitespace and quote escapes while preserving LaTeX backslashes
            temp_marker = "___BACKSLASH_ESCAPED___"
            s = raw_md.replace('\\\\', temp_marker)
            s = s.replace('\\"', '"')
            s = s.replace('\\n', '\n')
            s = s.replace('\\r', '\r')
            s = s.replace('\\t', '\t')
            s = s.replace(temp_marker, '\\')
            markdown_text = s
        else:
            markdown_text = resp_text

        return figures, markdown_text

    def clean_detailed_alt_text(raw_text, default_label="Detailed document figure"):
        """Cleans and formats a comprehensive, detailed alt text for accessibility."""
        if not raw_text or not raw_text.strip():
            return default_label
        # Remove leading redundant 'Figure X:' or 'Fig X.' prefixes
        text = re.sub(r'^(Figure|Fig\.?|Diagram|Image)\s*\d*[:.-]?\s*', '', raw_text.strip(), flags=re.IGNORECASE)
        # Convert LaTeX inline math $var$ to plain readable characters var
        text = re.sub(r'\$([^$]+)\$', r'\1', text)
        # Remove formatting symbols that break markdown image brackets or docPr XML attributes
        text = re.sub(r'[#*_`\[\]"]', '', text)
        # Normalize newlines and whitespace into single spaces so image tag doesn't break
        text = re.sub(r'\s+', ' ', text).strip()
        return text if text else default_label

    def crop_and_save_figures(figures, files_meta, target_dir, dpi=200):
        """Crops detected bounding boxes from source files and saves them to target_dir."""
        extracted = []
        pdf_cache = {}

        for fig_idx, fig in enumerate(figures):
            fig_id = fig.get("id") or f"figure_{fig_idx + 1}.png"
            if not fig_id.lower().endswith(('.png', '.jpg', '.jpeg')):
                fig_id = f"{fig_id}.png"

            file_idx = fig.get("file_index", 0)
            page_idx = fig.get("page_index", 0)
            caption = fig.get("caption", f"Figure {fig_idx + 1}")
            alt_text = clean_detailed_alt_text(fig.get("alt_text") or caption, f"Detailed illustration of Figure {fig_idx + 1}")
            box = fig.get("box_2d", [])

            if file_idx < 0 or file_idx >= len(files_meta):
                file_idx = 0
            file_info = files_meta[file_idx]

            try:
                # Render/get the page image
                if file_info["type"] == "pdf":
                    if file_idx not in pdf_cache:
                        pdf_cache[file_idx] = pymupdf.open(stream=file_info["bytes"], filetype="pdf")
                    doc = pdf_cache[file_idx]
                    safe_page_idx = min(max(0, page_idx), len(doc) - 1)
                    page = doc[safe_page_idx]
                    pix = page.get_pixmap(dpi=dpi)
                    source_img = Image.open(io.BytesIO(pix.tobytes("png")))
                    source_label = f"{file_info['name']} (Page {safe_page_idx + 1})"
                else:
                    source_img = Image.open(io.BytesIO(file_info["bytes"]))
                    source_label = file_info['name']

                # Crop using bounding box if available
                img_w, img_h = source_img.size
                if isinstance(box, (list, tuple)) and len(box) == 4:
                    ymin = max(0, int(box[0] / 1000.0 * img_h))
                    xmin = max(0, int(box[1] / 1000.0 * img_w))
                    ymax = min(img_h, int(box[2] / 1000.0 * img_h))
                    xmax = min(img_w, int(box[3] / 1000.0 * img_w))

                    # Ensure valid crop boundaries
                    if (xmax - xmin) > 10 and (ymax - ymin) > 10:
                        cropped_img = source_img.crop((xmin, ymin, xmax, ymax))
                    else:
                        cropped_img = source_img
                else:
                    cropped_img = source_img

                # Save cropped image into target directory
                save_path = os.path.join(target_dir, fig_id)
                cropped_img.save(save_path, format="PNG")

                extracted.append({
                    "id": fig_id,
                    "alt_text": alt_text,
                    "caption": caption,
                    "image": cropped_img,
                    "path": save_path,
                    "source": source_label,
                    "dimensions": f"{cropped_img.width}x{cropped_img.height}"
                })
            except Exception as crop_err:
                st.warning(f"Could not crop image {fig_id}: {crop_err}")

        # Fallback 1: If no figures were detected, attempt to extract embedded raster images from PDF files
        if len(extracted) == 0:
            for file_idx, file_info in enumerate(files_meta):
                if file_info["type"] == "pdf":
                    try:
                        if file_idx not in pdf_cache:
                            pdf_cache[file_idx] = pymupdf.open(stream=file_info["bytes"], filetype="pdf")
                        doc = pdf_cache[file_idx]
                        for page_idx, page in enumerate(doc):
                            image_list = page.get_images(full=True)
                            for img_info in image_list:
                                xref = img_info[0]
                                base_image = doc.extract_image(xref)
                                if base_image["width"] > 50 and base_image["height"] > 50:
                                    fig_id = f"figure_{len(extracted) + 1}.png"
                                    save_path = os.path.join(target_dir, fig_id)
                                    with open(save_path, "wb") as f_img:
                                        f_img.write(base_image["image"])
                                    pil_img = Image.open(io.BytesIO(base_image["image"]))
                                    extracted.append({
                                        "id": fig_id,
                                        "alt_text": f"Extracted figure {len(extracted) + 1} from page {page_idx + 1}",
                                        "caption": f"Extracted Figure {len(extracted) + 1}",
                                        "image": pil_img,
                                        "path": save_path,
                                        "source": f"{file_info['name']} (Page {page_idx + 1})",
                                        "dimensions": f"{pil_img.width}x{pil_img.height}"
                                    })
                    except Exception:
                        pass

        # Fallback 2: check if uploaded standalone images were referenced but not cropped
        saved_ids = {e["id"] for e in extracted}
        for f_info in files_meta:
            if f_info["name"] not in saved_ids:
                try:
                    direct_path = os.path.join(target_dir, f_info["name"])
                    with open(direct_path, "wb") as df:
                        df.write(f_info["bytes"])
                except Exception:
                    pass

        return extracted

    def create_zip_archive(extracted_figures):
        """Creates an in-memory zip archive of all extracted figures."""
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for fig in extracted_figures:
                if os.path.exists(fig["path"]):
                    zip_file.write(fig["path"], arcname=fig["id"])
            # Add manifest
            manifest = "\n".join([
                f"{f['id']}: Alt Text: \"{f.get('alt_text', f['caption'])}\" | Caption: \"{f['caption']}\" (Source: {f['source']}, {f['dimensions']})"
                for f in extracted_figures
            ])
            zip_file.writestr("images_manifest.txt", manifest)
        zip_buffer.seek(0)
        return zip_buffer.getvalue()

    def optimize_markdown_for_docx(markdown_text, extracted_figures, target_dir):
        r"""
        Optimizes Markdown content for Word (.docx) compilation via Pandoc:
        - Normalizes LaTeX display math (\[ \] -> $$) and inline math (\( \) -> $)
        - Trims whitespace adjacent to inline math dollar signs ($ var $ -> $var$) for Pandoc
        - Ensures display equations ($$...$$) have clean empty line padding
        - Ensures proper spacing around pipe tables, headers, blockquotes, and lists
        - Synchronizes all extracted figures into the Markdown text with detailed alt text
        - Resolves and aliases any referenced images so Pandoc can find them
        """
        if not markdown_text:
            return ""

        md = markdown_text

        # 1. Normalize LaTeX math syntax
        md = re.sub(r'\\\[([\s\S]*?)\\\]', r'$$\1$$', md)
        md = re.sub(r'\\\(([\s\S]*?)\\\)', r'$\1$', md)

        # 2. Trim whitespace immediately inside inline math delimiters ($ x $ -> $x$)
        md = re.sub(r'(?<!\$)\$\s+([^\$\n]+?)\s+\$(?!\$)', r'$\1$', md)
        md = re.sub(r'(?<!\$)\$\s+([^\$\n]+?)\$(?!\$)', r'$\1$', md)
        md = re.sub(r'(?<!\$)\$([^\$\n]+?)\s+\$(?!\$)', r'$\1$', md)

        # 3. Ensure display math ($$...$$) is isolated with clean empty lines
        md = re.sub(r'([^\n])\n(\$\$[\s\S]*?\$\$)', r'\1\n\n\2', md)
        md = re.sub(r'(\$\$[\s\S]*?\$\$)\n([^\n])', r'\1\n\n\2', md)

        # 4. Ensure pipe tables have clean spacing before and after the table block
        md = re.sub(r'([^\n|])\n(\|[^\n]+\|)', r'\1\n\n\2', md)
        md = re.sub(r'(\|[^\n]+\|)\n([^\n|])', r'\1\n\n\2', md)

        # 5. Ensure headings have clean line breaks
        md = re.sub(r'([^\n])\n(#{1,6}\s+.*)', r'\1\n\n\2', md)

        # 6. Ensure image blocks have clean line breaks around them
        md = re.sub(r'([^\n])\n(!\[.*?\]\(.*?\))', r'\1\n\n\2', md)
        md = re.sub(r'(!\[.*?\]\(.*?\))\n([^\n])', r'\1\n\n\2', md)

        # 7. Collapse excessive consecutive blank lines (more than 2 newlines -> 2)
        md = re.sub(r'\n{3,}', '\n\n', md)

        # 8. Synchronize all extracted figures: if any figure is not in the markdown, append it with detailed alt text
        alt_map = {f["id"]: f.get("alt_text", "Detailed document figure") for f in extracted_figures}
        for fig in extracted_figures:
            fig_id = fig.get("id", "")
            if fig_id and fig_id not in md:
                alt = fig.get("alt_text") or fig.get("caption", "Figure")
                caption = fig.get("caption", alt)
                md += f"\n\n![{alt}]({fig_id})\n\n*{caption}*\n"

        # 9. Enforce detailed alt text on all existing markdown image tags
        def update_img_tag(match):
            current_alt = match.group(1).strip()
            src_name = match.group(2).strip()
            if src_name in alt_map:
                return f"![{alt_map[src_name]}]({src_name})"
            elif current_alt:
                cleaned = clean_detailed_alt_text(current_alt)
                return f"![{cleaned}]({src_name})"
            return match.group(0)

        md = re.sub(r'!\[([^\]]*)\]\(([^)]+)\)', update_img_tag, md)

        # 10. Resolve image references: ensure any file referenced in ![...](filename) exists on disk
        existing_files = set(os.listdir(target_dir))
        for m in re.finditer(r'!\[([^\]]*)\]\(([^)]+)\)', md):
            ref_file = m.group(2).strip()
            if ref_file not in existing_files and extracted_figures:
                src_path = extracted_figures[0]["path"]
                if os.path.exists(src_path):
                    dest_path = os.path.join(target_dir, ref_file)
                    try:
                        with open(src_path, "rb") as sf, open(dest_path, "wb") as df:
                            df.write(sf.read())
                        existing_files.add(ref_file)
                    except Exception:
                        pass

        return md

    def optimize_docx_formatting(docx_path, extracted_figures=None):
        r"""
        Applies comprehensive professional document styling to the compiled Word document:
        1. Configures standard 1-inch margins on all sections.
        2. Centers all figures and enforces responsive image scaling to fit within page margins.
        3. Injects accessibility Alt Text (descr) and titles into DrawingML OpenXML metadata.
        4. Sets keep_with_next on images and headings to eliminate orphaned lines/images.
        5. Optimizes tables: centers tables, enables tblHeader on header rows, and cantSplit on all rows.
        """
        if extracted_figures is None:
            extracted_figures = []

        try:
            doc = docx.Document(docx_path)
            max_img_width = Inches(6.0)

            # 1. Page Margins (1 inch on all sides)
            for section in doc.sections:
                section.top_margin = Inches(1.0)
                section.bottom_margin = Inches(1.0)
                section.left_margin = Inches(1.0)
                section.right_margin = Inches(1.0)
                page_width = section.page_width or Inches(8.5)
                content_width = page_width - section.left_margin - section.right_margin
                max_img_width = min(content_width, Inches(6.0))

            # 2. Heading formatting & orphan prevention
            for para in doc.paragraphs:
                if para.style and para.style.name and para.style.name.startswith('Heading'):
                    para.paragraph_format.keep_with_next = True
                    para.paragraph_format.space_before = Pt(12)
                    para.paragraph_format.space_after = Pt(4)

            # 3. Figure Scaling & Accessibility DrawingML metadata
            fig_list = list(extracted_figures)
            for i, shape in enumerate(doc.inline_shapes):
                try:
                    # Scale down oversized images to fit nicely within printable margins
                    if shape.width and shape.width > max_img_width:
                        aspect = float(shape.height) / float(shape.width)
                        shape.width = int(max_img_width)
                        shape.height = int(max_img_width * aspect)

                    # Ingest DrawingML accessibility metadata
                    doc_prs = shape._inline.xpath('.//wp:docPr')
                    if doc_prs:
                        doc_pr = doc_prs[0]
                        if i < len(fig_list):
                            target_alt = fig_list[i].get("alt_text") or "Detailed document figure"
                            target_title = fig_list[i].get("caption") or f"Figure {i + 1}"
                        elif extracted_figures:
                            target_alt = extracted_figures[0].get("alt_text") or "Detailed document figure"
                            target_title = extracted_figures[0].get("caption") or "Figure 1"
                        else:
                            target_alt = "Detailed document figure"
                            target_title = "Figure"

                        doc_pr.set('descr', target_alt)
                        doc_pr.set('title', target_title)
                except Exception:
                    pass

            # 4. Center-align image paragraphs and pair with captions
            for para in doc.paragraphs:
                drawings = para._p.xpath('.//w:drawing')
                if drawings:
                    para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                    para.paragraph_format.keep_with_next = True
                    para.paragraph_format.space_before = Pt(8)
                    para.paragraph_format.space_after = Pt(2)
                elif para.text and para.text.strip().lower().startswith(('figure ', 'fig.', 'fig ')):
                    para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                    para.paragraph_format.space_before = Pt(2)
                    para.paragraph_format.space_after = Pt(10)

            # 5. Table Optimization (Centering, Header repeats, cantSplit rows)
            for table in doc.tables:
                try:
                    table.alignment = WD_TABLE_ALIGNMENT.CENTER
                    for row_idx, row in enumerate(table.rows):
                        trPr = row._tr.get_or_add_trPr()
                        if not trPr.xpath('./w:cantSplit'):
                            trPr.append(parse_xml(r'<w:cantSplit xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'))
                        if row_idx == 0:
                            if not trPr.xpath('./w:tblHeader'):
                                trPr.append(parse_xml(r'<w:tblHeader xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'))
                except Exception:
                    pass

            doc.save(docx_path)
        except Exception:
            pass

    # --- Process Pipeline ---
    if st.button("🚀 Process and Convert", type="primary"):
        if not final_key:
            st.error("Please provide a valid Gemini API Key to proceed. Follow the instructions above to generate one.")
        elif not uploaded_files:
            st.warning("Please upload at least one PNG or PDF file.")
        elif not output_name.strip():
            st.error("Please enter a valid output document name.")
        else:
            with st.spinner("Processing... analyzing layout, extracting figures with Gemini AI, and compiling Word document..."):
                try:
                    # 1. Initialize Gemini Client
                    client = genai.Client(api_key=final_key)

                    # 2. Inspect uploaded files & build manifest
                    files_meta, file_parts, manifest_text = prepare_file_manifest(uploaded_files)

                    # 3. Formulate Prompt
                    prompt = f"""You are an expert document transcription, OCR, and visual layout analysis AI.
You are analyzing a series of sequential document files (images / PDF pages):
{manifest_text}

CORE INSTRUCTIONS:
1. ACCURATE TRANSCRIPTION & STRUCTURE:
   - Extract all text, tables, and structural elements from the provided files and output them cleanly formatted in standard Markdown.
   - Maintain the logical reading order, headings hierarchy (# H1, ## H2, ### H3), section numbering, bullet points, and numbered lists.
   - Do NOT include conversational filler like 'Here is your markdown', 'Certainly', or conversational preamble/conclusion. Output only the requested JSON format.

2. TABLES & DATA:
   - Convert all tabular data into valid standard Markdown tables with proper header rows and delimiter alignment lines (|---|---|).
   - Ensure cell contents, column headers, and structural alignments match the original document.

3. MATHEMATICS & EQUATIONS:
   - Enclose all equations and inline mathematical expressions in single dollar signs like $e=mc^2$.
   - Enclose standalone or multi-line block equations in double dollar signs $$...$$.

4. FIGURE & IMAGE EXTRACTION (CRITICAL INSTRUCTION FOR IMAGES):
   - Thoroughly detect ALL visual figures, diagrams, charts, plots, graphs, drawings, schematics, illustrations, and photos across every page and file.
   - For each detected visual figure or diagram:
     * Assign a unique sequential identifier: "figure_1.png", "figure_2.png", "figure_3.png", etc.
     * Set "file_index": 0-based integer index of the file containing this figure.
     * Set "page_index": 0-based integer index of the page in that file (for PNG/images, page_index is always 0; for PDF, 0 is the 1st page, 1 is the 2nd page, etc.).
     * Set "box_2d": [ymin, xmin, ymax, xmax] coordinates normalized to 0-1000 on that specific page/image. Tightly bound the entire graphic (and its title/caption/legend if visually part of the figure).
     * Set "alt_text": A detailed, comprehensive alternative text description (typically 1-3 detailed, descriptive sentences) explaining the visual elements, diagram structure, labels, variables, axes, data trends, components, connections, or layout shown in the figure for accessibility and screen readers (e.g., "A detailed schematic diagram of a cantilever beam with fixed boundary support at the left end, subjected to a concentrated downward point load P applied at the free right end of length L, depicting coordinate axes x and y and the resulting deflection curve v(x)").
     * Set "caption": Full descriptive caption or title of the figure from the document.
   - IN THE TRANSCRIBED MARKDOWN: Whenever you reference, describe, or analyze a specific image in your report, you MUST embed the figure inline using standard Markdown image syntax with its detailed alt text:
     ![Detailed alt text describing visual elements](figure_X.png)
     Followed immediately by any explanatory caption or title from the document.

RESPONSE FORMAT:
You MUST respond with a valid JSON object matching this schema:
{{
  "figures": [
    {{
      "id": "figure_1.png",
      "file_index": 0,
      "page_index": 0,
      "alt_text": "Detailed, comprehensive description of the visual figure, components, labels, and trends",
      "caption": "Full descriptive caption",
      "box_2d": [ymin, xmin, ymax, xmax]
    }}
  ],
  "markdown": "... full transcribed document in clean standard Markdown format with inline ![detailed alt text](figure_X.png) tags ..."
}}
Do NOT output conversational filler or Markdown code blocks around the JSON. Return only the JSON object.
"""

                    contents = [prompt] + file_parts

                    # 4. Generate Content via Gemini with robust document safety settings
                    safety_settings = [
                        types.SafetySetting(category="HARM_CATEGORY_HATE_SPEECH", threshold="BLOCK_ONLY_HIGH"),
                        types.SafetySetting(category="HARM_CATEGORY_DANGEROUS_CONTENT", threshold="BLOCK_ONLY_HIGH"),
                        types.SafetySetting(category="HARM_CATEGORY_HARASSMENT", threshold="BLOCK_ONLY_HIGH"),
                        types.SafetySetting(category="HARM_CATEGORY_SEXUALLY_EXPLICIT", threshold="BLOCK_ONLY_HIGH"),
                    ]

                    response = client.models.generate_content(
                        model=model_choice,
                        contents=contents,
                        config=types.GenerateContentConfig(
                            response_mime_type="application/json",
                            safety_settings=safety_settings
                        )
                    )

                    # 5. Parse Gemini output with robust LaTeX and JSON recovery
                    resp_text = response.text.strip()
                    figures_data, markdown_text = parse_gemini_document_response(resp_text)

                    # 6. Setup temporary directory for Pandoc conversion & image cropping
                    with tempfile.TemporaryDirectory() as tmpdir:
                        # Crop and save all detected figures into tmpdir
                        extracted_figures = crop_and_save_figures(
                            figures_data, 
                            files_meta, 
                            tmpdir, 
                            dpi=render_dpi
                        )

                        # Optimize Markdown for Word (math, tables, figures, image tags)
                        optimized_markdown = optimize_markdown_for_docx(
                            markdown_text, 
                            extracted_figures, 
                            tmpdir
                        )

                        # Write optimized Markdown file
                        md_path = os.path.join(tmpdir, "document.md")
                        docx_path = os.path.join(tmpdir, "document.docx")

                        with open(md_path, "w", encoding="utf-8") as md_file:
                            md_file.write(optimized_markdown)

                        # Convert to docx via Pandoc with rich extensions and embedded resource resolution
                        pandoc_args = [
                            '--from=markdown+tex_math_dollars+pipe_tables+raw_html+grid_tables+auto_identifiers',
                            f'--resource-path={tmpdir}',
                            '--standalone'
                        ]

                        pypandoc.convert_file(
                            md_path, 
                            'docx', 
                            outputfile=docx_path,
                            extra_args=pandoc_args
                        )

                        # Optimize Word document styling, tables, image scaling, and accessibility metadata
                        optimize_docx_formatting(docx_path, extracted_figures)

                        # Read compiled docx back into memory
                        with open(docx_path, "rb") as docx_file:
                            docx_bytes = docx_file.read()

                        # Create ZIP archive of extracted images if any exist
                        zip_bytes = create_zip_archive(extracted_figures) if extracted_figures else None

                    # 7. Provide Results & Downloads to the User
                    st.success(f"🎉 Conversion Complete! Extracted {len(extracted_figures)} image(s)/figure(s) and compiled Word document.")

                    clean_name = output_name.strip().replace(".docx", "").replace(".md", "").replace(".zip", "")

                    # Download Buttons
                    if zip_bytes:
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            st.download_button(
                                label="📥 Download Word Doc (.docx)",
                                data=docx_bytes,
                                file_name=f"{clean_name}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True
                            )
                        with col2:
                            st.download_button(
                                label="📄 Download Markdown (.md)",
                                data=optimized_markdown,
                                file_name=f"{clean_name}.md",
                                mime="text/markdown",
                                use_container_width=True
                            )
                        with col3:
                            st.download_button(
                                label="📦 Download Images (.zip)",
                                data=zip_bytes,
                                file_name=f"{clean_name}_images.zip",
                                mime="application/zip",
                                use_container_width=True
                            )
                    else:
                        col1, col2 = st.columns(2)
                        with col1:
                            st.download_button(
                                label="📥 Download Word Doc (.docx)",
                                data=docx_bytes,
                                file_name=f"{clean_name}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True
                            )
                        with col2:
                            st.download_button(
                                label="📄 Download Markdown (.md)",
                                data=optimized_markdown,
                                file_name=f"{clean_name}.md",
                                mime="text/markdown",
                                use_container_width=True
                            )

                except Exception as e:
                    st.error(f"An unexpected error occurred during processing:\n{e}")
