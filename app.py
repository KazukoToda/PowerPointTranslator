import streamlit as st
import os
from dotenv import load_dotenv
import tempfile
from pptx_processor import extract_text_from_pptx, create_translated_pptx
from translator import translate_text

# Load environment variables
load_dotenv()

# Set page configuration
st.set_page_config(
    page_title="PowerPoint Translator - 日本語から英語へ",
    page_icon="📊",
    layout="centered"
)

# App title and description
st.title("PowerPoint Translator")
st.markdown("日本語から英語へ PowerPointファイルを翻訳")

# Function to process the uploaded file
def process_pptx(uploaded_file):
    # Create a temporary file
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as tmp_file:
        tmp_file.write(uploaded_file.getvalue())
        input_path = tmp_file.name
    
    try:
        # Extract text from PPTX
        with st.spinner('PowerPointファイルからテキストを抽出中...'):
            slides_data = extract_text_from_pptx(input_path)
        
        # Translate text
        with st.spinner('テキストを翻訳中...'):
            translated_slides = []
            for slide in slides_data:
                translated_elements = []
                for element in slide:
                    text, shape_info = element
                    if text and text.strip():
                        translated_text = translate_text(text)
                    else:
                        translated_text = text
                    translated_elements.append((translated_text, shape_info))
                translated_slides.append(translated_elements)
        
        # Create translated PPTX
        with st.spinner('翻訳されたPowerPointファイルを作成中...'):
            output_path = input_path.replace('.pptx', '_translated.pptx')
            create_translated_pptx(input_path, output_path, translated_slides)
        
        return output_path
    except Exception as e:
        st.error(f"エラーが発生しました: {str(e)}")
        raise e
    finally:
        # Clean up the input temporary file
        if os.path.exists(input_path):
            os.remove(input_path)

# Custom dictionary upload section (optional)
st.markdown("### カスタム辞書 (オプション)")
custom_dict_file = st.file_uploader("カスタム辞書ファイル (CSVまたはExcel)", type=["csv", "xlsx"], help="カンマ区切りの「原文,訳文」形式のファイル")

# Main file upload section
st.markdown("### PowerPointファイルのアップロード")
uploaded_file = st.file_uploader("PowerPointファイル", type=["pptx"])

if uploaded_file is not None:
    st.write("ファイル名:", uploaded_file.name)
    
    # Process button
    if st.button("翻訳開始"):
        try:
            output_file_path = process_pptx(uploaded_file)
            
            # Display success message
            st.success("翻訳が完了しました！")
            
            # Provide download link
            with open(output_file_path, "rb") as file:
                btn = st.download_button(
                    label="翻訳ファイルをダウンロード",
                    data=file,
                    file_name=f"translated_{uploaded_file.name}",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
            
            # Clean up the output temporary file
            if os.path.exists(output_file_path):
                os.remove(output_file_path)
                
        except Exception as e:
            st.error(f"ファイルの処理中にエラーが発生しました: {str(e)}")

# Footer
st.markdown("---")
st.markdown("Powered by Azure Translator API")