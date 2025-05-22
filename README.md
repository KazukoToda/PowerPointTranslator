# PowerPoint Translator

A web application that allows users to upload PowerPoint (PPTX) files, translates Japanese text to English using Azure Translator API, and provides the translated PowerPoint file for download.

## Features

- Upload PowerPoint (PPTX) files
- Translate Japanese text to English using Azure Translator API
- Download translated PowerPoint files
- User-friendly interface with Streamlit
- Optional custom dictionary support

## Prerequisites

- Python 3.8 or higher
- Azure Translator API key (with subscription)

## Installation

1. Clone the repository:

```bash
git clone https://github.com/KazukoToda/PowerPointTranslator.git
cd PowerPointTranslator
```

2. Create a virtual environment and activate it:

```bash
python -m venv venv
source venv/bin/activate  # On Windows, use: venv\Scripts\activate
```

3. Install dependencies:

```bash
pip install -r requirements.txt
```

4. Create a `.env` file based on the example:

```bash
cp .env.example .env
```

5. Add your Azure Translator API key and region to the `.env` file:

```
AZURE_TRANSLATOR_KEY=your_azure_translator_key_here
AZURE_TRANSLATOR_REGION=your_azure_region_here
```

## Usage

1. Start the Streamlit app:

```bash
streamlit run app.py
```

2. Open your web browser and navigate to the URL displayed in the terminal (typically `http://localhost:8501`)

3. Upload a PowerPoint file using the web interface

4. Optionally upload a custom dictionary file (CSV format with "source,target" pairs)

5. Click "翻訳開始" (Start Translation)

6. Wait for the translation to complete

7. Download the translated PowerPoint file

## Limitations

- Only supports PPTX files (not PPT)
- Simple text translation (complex formatting may be affected)
- Images and other elements are preserved but not modified
- Some complex slide layouts may not be preserved perfectly