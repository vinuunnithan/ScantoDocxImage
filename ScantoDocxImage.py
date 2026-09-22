import os
import tempfile
import pypandoc
import streamlit as st
from google import genai

# Setup Page Configuration
st.set_page_config(page_title="Document to Word Converter", page_icon="📝", layout="centered")


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
    st.subheader("✍️ Digital Acknowledgement")
    # Interactive signature text entry
    user_signature = st.text_input(
        "Type your full name to accept these terms:", 
        placeholder="First and Last Name",
        help="Entering your name acts as a digital signature acknowledging the risks listed above."
    ).strip()

    # Track if the disclaimer has been signed
    is_signed = len(user_signature) > 0

#st.title("<h1 style='font-size: 32px;'> 📝 PNG/PDF→Markdown→Word Converter")

st.markdown("<h1 style='font-size: 32px;'>📝 PNG/PDF→Markdown→Word Converter</h1>", unsafe_allow_html=True)
st.write("Upload your PNG images or PDF files, convert them to clean Markdown via Gemini AI, and download an **editable Word Document**. Developed by Dr. Vinu Unnikrishnan in collaboration with **Gemini AI**.")

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
        2. **Create Key:** Click the blue **"Get API key"** button in the bottom-left corner of the dashboard (🔑).
        3. **Copy Key:** Choose **"Create API key"**, select or create a project, and copy your generated key string.
        4. **Paste Below:** Paste that key string into the field below to unlock the file processor.
        """
    )
    st.write("---")

    # --- API Key Input Field ---
    env_key = os.environ.get("GEMINI_API_KEY", "")

    if not env_key:
        api_key_input = st.text_input(
            "Enter your Gemini API Key:", 
            type="password", 
            help="Paste the key you generated from https://google.com"
        )
        final_key = api_key_input.strip()
    else:
        final_key = env_key
        st.info("🤖 Gemini API Key loaded securely from Cloud Environment.")

    # --- File Upload Section (Accepts PNG and PDF) ---
    uploaded_files = st.file_uploader(
        "Drag and drop or browse PNG/PDF files:", 
        type=["png", "pdf"], 
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

    # --- Process Pipeline ---
    if st.button("🚀 Process and Convert", type="primary"):
        if not final_key:
            st.error("Please provide a valid Gemini API Key to proceed. Follow the instructions above to generate one.")
        elif not uploaded_files:
            st.warning("Please upload at least one PNG or PDF file.")
        elif not output_name.strip():
            st.error("Please enter a valid output document name.")
        else:
            with st.spinner("Processing... calling Gemini AI and compiling Word document..."):
                try:
                    # 1. Initialize Gemini Client
                    client = genai.Client(api_key=final_key)
                    
                    contents = ["Extract all text, tables, and structural elements from this image and output them cleanly formatted in standard Markdown. Do not include conversational filler like 'Here is your markdown'. Enclose all equations in single dollar signs like $e=mc^2$ You are analyzing a series of sequential PNG images. CRITICAL INSTRUCTION FOR IMAGES:. whenever you reference, describe, or analyze a specific image in your report, you MUST embed the figure inline using standard Markdown image syntax." ]
                    
                    # Load file bytes directly into memory buffers with dynamic MIME handling
                    for uploaded_file in uploaded_files:
                        file_bytes = uploaded_file.read()
                        
                        # VERIFICATION LOGIC: Determine exact MIME type based on file suffix
                        if uploaded_file.name.lower().endswith('.pdf'):
                            mime_type = 'application/pdf'
                        else:
                            mime_type = 'image/png'
                            
                        contents.append(
                            genai.types.Part.from_bytes(data=file_bytes, mime_type=mime_type)
                        )

                    # 2. Generate Content via Gemini
                    response = client.models.generate_content(
                        model='gemini-3.6-flash',
                        contents=contents
                    )
                    markdown_text = response.text

                    # 3. Handle Pandoc Conversion using secure temporary cloud directories
                    with tempfile.TemporaryDirectory() as tmpdir:
                        md_path = os.path.join(tmpdir, "temp.md")
                        docx_path = os.path.join(tmpdir, "temp.docx")
                        
                        # Write markdown to a temporary cloud file
                        with open(md_path, "w", encoding="utf-8") as md_file:
                            md_file.write(markdown_text)
                        
                        # Convert to docx via Pandoc
                        pypandoc.convert_file(md_path, 'docx', outputfile=docx_path)
                        
                        # Read the compiled docx file back into memory
                        with open(docx_path, "rb") as docx_file:
                            docx_bytes = docx_file.read()

                    # 4. Provide Downloads to the User
                    st.success("🎉 Conversion Complete!")
                    
                    # Format final filenames cleanly
                    clean_name = output_name.strip().replace(".docx", "").replace(".md", "")
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        st.download_button(
                            label="📥 Download Word Document (.docx)",
                            data=docx_bytes,
                            file_name=f"{clean_name}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )
                    with col2:
                        st.download_button(
                            label="📄 Download Raw Markdown (.md)",
                            data=markdown_text,
                            file_name=f"{clean_name}.md",
                            mime="text/markdown",
                            use_container_width=True
                        )

                except Exception as e:
                    st.error(f"An unexpected error occurred during processing:\n{e}")
