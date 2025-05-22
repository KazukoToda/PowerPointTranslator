const fs = require('fs');
const path = require('path');
const axios = require('axios');
const pptxExtractor = require('pptx-extractor');
const officegen = require('officegen');

/**
 * Translates Japanese text to English using Azure Translator API
 * @param {string} text - Text to translate
 * @returns {Promise<string>} - Translated text
 */
async function translateText(text) {
  if (!text || text.trim() === '') {
    return text;
  }

  const apiKey = process.env.AZURE_TRANSLATOR_KEY;
  const region = process.env.AZURE_TRANSLATOR_REGION;
  
  if (!apiKey || !region) {
    throw new Error('Azure Translator API key or region not configured');
  }

  const endpoint = 'https://api.cognitive.microsofttranslator.com';
  const url = `${endpoint}/translate?api-version=3.0&from=ja&to=en`;

  try {
    const response = await axios({
      method: 'post',
      url: url,
      headers: {
        'Ocp-Apim-Subscription-Key': apiKey,
        'Ocp-Apim-Subscription-Region': region,
        'Content-Type': 'application/json'
      },
      data: [{ text }]
    });

    if (response.data && response.data.length > 0 && 
        response.data[0].translations && 
        response.data[0].translations.length > 0) {
      return response.data[0].translations[0].text;
    } else {
      throw new Error('Translation API returned unexpected response');
    }
  } catch (error) {
    console.error('Translation error:', error.response?.data || error.message);
    throw new Error(`Failed to translate text: ${error.message}`);
  }
}

/**
 * Translates content of a PowerPoint file from Japanese to English
 * @param {string} filePath - Path to the PowerPoint file
 * @returns {Promise<string>} - Path to the translated file
 */
async function translatePowerPoint(filePath) {
  try {
    // Extract content from PPTX
    const extractedData = await pptxExtractor.extract(filePath);
    
    // Create a new PowerPoint file
    const pptx = officegen('pptx');

    // Process each slide
    for (const slide of extractedData.slides) {
      // Create a new slide
      const newSlide = pptx.makeNewSlide();

      // Get all text elements from the slide
      const textElements = slide.textElements || [];
      
      // Translate and add text elements to the new slide
      for (const textElement of textElements) {
        if (textElement.text && textElement.text.trim()) {
          const translatedText = await translateText(textElement.text);
          
          // Add text to the slide (simplified - adjust coordinates and formatting as needed)
          newSlide.addText(translatedText, {
            x: textElement.position?.x || 0,
            y: textElement.position?.y || 0,
            font_size: textElement.fontSize || 18,
            color: textElement.color || '000000'
          });
        }
      }
      
      // Handle images if present (copying from original)
      if (slide.images && slide.images.length > 0) {
        for (const image of slide.images) {
          // Add images to the new slide (simplified)
          if (image.path) {
            newSlide.addImage({ 
              path: image.path,
              x: image.position?.x || 0,
              y: image.position?.y || 0,
              cx: image.size?.width || 300,
              cy: image.size?.height || 200
            });
          }
        }
      }
    }

    // Generate the translated PowerPoint file
    const fileName = path.basename(filePath, '.pptx');
    const translatedFilePath = path.join(
      path.dirname(filePath),
      `${fileName}_translated.pptx`
    );

    return new Promise((resolve, reject) => {
      const outputStream = fs.createWriteStream(translatedFilePath);
      
      // Handle events
      pptx.on('error', (err) => {
        console.error('Error generating PPTX:', err);
        reject(err);
      });
      
      outputStream.on('error', (err) => {
        console.error('Error writing file:', err);
        reject(err);
      });
      
      outputStream.on('close', () => {
        resolve(translatedFilePath);
      });
      
      // Generate the PPTX file
      pptx.generate(outputStream);
    });
  } catch (error) {
    console.error('Error processing PowerPoint file:', error);
    throw new Error(`Failed to translate PowerPoint: ${error.message}`);
  }
}

module.exports = {
  translatePowerPoint,
  translateText
};