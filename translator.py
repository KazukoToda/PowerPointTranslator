import os
import json
import requests
import uuid
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

def translate_text(text, source_lang='ja', target_lang='en', custom_dict=None):
    """
    Translate text using Azure Translator API
    
    Args:
        text (str): Text to translate
        source_lang (str): Source language code (default: 'ja')
        target_lang (str): Target language code (default: 'en')
        custom_dict (dict): Custom dictionary for terminology (default: None)
    
    Returns:
        str: Translated text
    """
    # If text is empty, return as is
    if not text or text.strip() == "":
        return text
    
    # Apply custom dictionary if provided
    if custom_dict:
        for key, value in custom_dict.items():
            if key in text:
                text = text.replace(key, value)
    
    # Get API key and region from environment variables
    api_key = os.getenv("AZURE_TRANSLATOR_KEY")
    region = os.getenv("AZURE_TRANSLATOR_REGION")
    
    if not api_key or not region:
        raise ValueError("Azure Translator API key or region not configured")
    
    # API endpoint and parameters
    endpoint = "https://api.cognitive.microsofttranslator.com"
    path = "/translate"
    constructed_url = endpoint + path
    
    params = {
        "api-version": "3.0",
        "from": source_lang,
        "to": target_lang
    }
    
    headers = {
        "Ocp-Apim-Subscription-Key": api_key,
        "Ocp-Apim-Subscription-Region": region,
        "Content-type": "application/json",
        "X-ClientTraceId": str(uuid.uuid4())
    }
    
    # Prepare body for the request
    body = [{
        "text": text
    }]
    
    # Make the API request
    try:
        response = requests.post(constructed_url, params=params, headers=headers, json=body)
        response.raise_for_status()  # Raise exception for HTTP errors
        
        # Parse the response
        result = response.json()
        
        # Extract translated text
        if result and len(result) > 0 and "translations" in result[0]:
            translated_text = result[0]["translations"][0]["text"]
            return translated_text
        else:
            raise ValueError("Translation API returned unexpected response")
            
    except requests.exceptions.RequestException as e:
        raise Exception(f"Translation request failed: {str(e)}")
    except (KeyError, IndexError, json.JSONDecodeError) as e:
        raise Exception(f"Error parsing translation response: {str(e)}")

def load_custom_dictionary(file_path):
    """
    Load custom dictionary from CSV or Excel file
    
    Args:
        file_path (str): Path to the CSV or Excel file
    
    Returns:
        dict: Dictionary with source terms as keys and translations as values
    """
    custom_dict = {}
    
    # Implement dictionary loading based on file type
    # This is a placeholder for the actual implementation
    
    return custom_dict