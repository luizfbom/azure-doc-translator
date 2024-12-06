import os
import requests
from utils.config import Config
import streamlit as st

def translate_pptx(text_list, target_language):
    """
    Translates a list of texts in a single API call to Azure Translator.
    
    Args:
        text_list (list): List of texts to translate
        target_language (str): Target language code
        
    Returns:
        dict: Dictionary mapping original texts to their translations
    """
    headers = {
        'Ocp-Apim-Subscription-Key': Config.SUBSCRIPTION_KEY,
        'Ocp-Apim-Subscription-Region': Config.AZURE_LOCATION,
        'Content-type': 'application/json',
        'X-ClientTraceId': str(os.urandom(16))
    }
    
    constructed_url = Config.ENDPOINT + '/translate'
    
    params = {
        'api-version': '3.0',
        'to': target_language
    }
    
    # Prepare the body with all texts at once
    body = [{'text': text} for text in text_list if text.strip()]
    
    translations_dict = {}
    
    try:
        # Single API call for all texts
        response = requests.post(
            constructed_url,
            params=params,
            headers=headers,
            json=body
        )
        response.raise_for_status()
        
        # Process all translations
        translations = response.json()
        
        # Collect translation results
        for original_text, translation_result in zip(text_list, translations):
            if original_text.strip():
                translated_text = translation_result['translations'][0]['text']
                translations_dict[original_text] = translated_text
                
    except requests.exceptions.RequestException as e:
        st.error(f"Error during translation: {e}")
        translations_dict = {text: "" for text in text_list if text.strip()}
    
    return translations_dict

