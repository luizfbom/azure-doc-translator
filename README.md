# Azure PowerPoint Translator

A Python application that translates Microsoft PowerPoint presentations using Azure's Translator API.

## Setup

1. Clone the repository
2. Install the dependencies: `pip install -r requirements.txt`
3. Create a `.env` file with the following variables:
   - SUBSCRIPTION_KEY
   - ENDPOINT
   - AZURE_LOCATION
4. Run the application: `streamlit run src/app.py`

## Supported Languages
Supports all languages available in Azure Translator service. Use standard language codes (e.g., 'es' for Spanish).
