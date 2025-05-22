# PowerPoint Translator

A web application that allows users to upload PowerPoint (PPTX) files, translates Japanese text to English using Azure Translator API, and provides the translated PowerPoint file for download.

## Features

- Upload PowerPoint (PPTX) files
- Translate Japanese text to English using Azure Translator API
- Download translated PowerPoint files
- User-friendly interface with drag-and-drop support

## Prerequisites

- Node.js (v14 or higher)
- Azure Translator API key (with subscription)

## Installation

1. Clone the repository:

```bash
git clone https://github.com/KazukoToda/PowerPointTranslator.git
cd PowerPointTranslator
```

2. Install dependencies:

```bash
npm install
```

3. Create a `.env` file based on the example:

```bash
cp .env.example .env
```

4. Add your Azure Translator API key and region to the `.env` file:

```
AZURE_TRANSLATOR_KEY=your_azure_translator_key_here
AZURE_TRANSLATOR_REGION=your_azure_region_here
PORT=3000
```

## Usage

1. Start the server:

```bash
npm start
```

2. Open your web browser and navigate to `http://localhost:3000`

3. Upload a PowerPoint file using the web interface

4. Wait for the translation to complete

5. Download the translated PowerPoint file

## Development

For development with automatic restarts, run:

```bash
npm run dev
```

## Limitations

- Only supports PPTX files (not PPT)
- Maximum file size is 50MB
- Simple text translation (formatting may be affected)
- Images and other elements are preserved but not modified