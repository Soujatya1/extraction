import streamlit as st
import pdfplumber
import pandas as pd
import docx
import json
import io
import os
import zipfile

def extract_tables_from_pdf(file_path):
    document_content = []
    
    with pdfplumber.open(file_path) as pdf:
        for page_num, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            if text.strip():
                document_content.append({
                    "content": text,
                    "page": page_num + 1,
                    "type": "text"
                })
            
            tables = page.extract_tables()
            for table_num, table in enumerate(tables):
                if table:
                    df = pd.DataFrame(table)
                    
                    if not df.empty:
                        headers = []
                        if len(df.columns) > 0:
                            if not pd.isna(df.iloc[0]).all() and not all(x is None for x in df.iloc[0]):
                                headers = [str(h).strip() if h is not None else f"Column_{i}" 
                                          for i, h in enumerate(df.iloc[0])]
                                df = df.iloc[1:]
                            else:
                                headers = [f"Column_{i}" for i in range(len(df.columns))]
                        
                        unique_headers = []
                        header_counts = {}
                        
                        for h in headers:
                            if h in header_counts:
                                header_counts[h] += 1
                                unique_headers.append(f"{h}_{header_counts[h]}")
                            else:
                                header_counts[h] = 0
                                unique_headers.append(h)
                        
                        df.columns = unique_headers
                    
                    document_content.append({
                        "page": page_num + 1,
                        "type": "table",
                        "table_number": table_num + 1,
                        "dataframe": df
                    })
    
    return document_content

def create_word_document_text_only(document_content):
    doc = docx.Document()
    doc.add_heading('PDF Text Content', 0)
    
    text_content = [item for item in document_content if item["type"] == "text"]
    
    for item in text_content:
        doc.add_heading(f'Page {item["page"]}', level=1)
        
        doc.add_paragraph(item["content"])
    
    docx_io = io.BytesIO()
    doc.save(docx_io)
    docx_io.seek(0)
    
    return docx_io

def create_json_tables(document_content):
    table_content = [item for item in document_content if item["type"] == "table"]
    
    tables_dict = {}
    
    for item in table_content:
        df = item["dataframe"]
        table_data = df.to_dict(orient="records")
        
        table_key = f"page_{item['page']}_table_{item['table_number']}"
        
        tables_dict[table_key] = {
            "page": item["page"],
            "table_number": item["table_number"],
            "data": table_data
        }
    
    json_string = json.dumps(tables_dict, indent=2)
    
    return json_string

def create_zip_archive(word_doc, json_data, base_filename):
    zip_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        zip_file.writestr(f"{base_filename}_text.docx", word_doc.getvalue())
        
        zip_file.writestr(f"{base_filename}_tables.json", json_data)
    
    zip_buffer.seek(0)
    
    return zip_buffer

st.title("PDF Extractor")

uploaded_file = st.file_uploader("Upload a PDF", type="pdf")

if uploaded_file:
    temp_dir = 'temp_pdfs'
    os.makedirs(temp_dir, exist_ok=True)
    
    file_path = f"{temp_dir}/{uploaded_file.name}"
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    
    with st.spinner("Extracting content from PDF..."):
        document_content = extract_tables_from_pdf(file_path)
    
    text_count = sum(1 for item in document_content if item["type"] == "text")
    table_count = sum(1 for item in document_content if item["type"] == "table")
    
    st.success(f"Successfully extracted {text_count} text sections and {table_count} tables from {uploaded_file.name}")
    
    with st.expander("Preview Extracted Content"):
        st.subheader("Text Content Preview")
        text_items = [item for item in document_content if item["type"] == "text"]
        if text_items:
            for i, item in enumerate(text_items[:2]):
                st.write(f"**Page {item['page']} - Text**")
                st.text(item["content"][:300] + ("..." if len(item["content"]) > 300 else ""))
                if i < min(1, len(text_items) - 1):
                    st.divider()
            
            if len(text_items) > 2:
                st.write(f"*...and {len(text_items) - 2} more text sections*")
        else:
            st.write("No text content found in the PDF.")
        
        st.subheader("Table Content Preview")
        table_items = [item for item in document_content if item["type"] == "table"]
        if table_items:
            for i, item in enumerate(table_items[:2]):
                st.write(f"**Page {item['page']} - Table {item['table_number']}**")
                df = item.get("dataframe")
                if df is not None and not df.empty:
                    st.dataframe(df)
                if i < min(1, len(table_items) - 1):
                    st.divider()
            
            if len(table_items) > 2:
                st.write(f"*...and {len(table_items) - 2} more tables*")
        else:
            st.write("No tables found in the PDF.")
    
    if document_content:
        word_doc = create_word_document_text_only(document_content)
        
        json_data = create_json_tables(document_content)
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.download_button(
                label="Download Text as Word",
                data=word_doc,
                file_name=f"{uploaded_file.name.split('.')[0]}_text.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        
        with col2:
            st.download_button(
                label="Download Tables as JSON",
                data=json_data,
                file_name=f"{uploaded_file.name.split('.')[0]}_tables.json",
                mime="application/json"
            )
        
        base_filename = uploaded_file.name.split('.')[0]
        zip_buffer = create_zip_archive(word_doc, json_data, base_filename)
        
        st.download_button(
            label="Download All (ZIP)",
            data=zip_buffer,
            file_name=f"{base_filename}_extraction.zip",
            mime="application/zip"
        )
    
    try:
        os.remove(file_path)
    except:
        pass
else:
    st.info("Please upload a PDF file to extract its content")
