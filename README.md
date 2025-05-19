# AzureAI
Azure AI Services Integration Sample

This repository demonstrates how to integrate various Azure AI services, including:
- Azure Language Service (for text analysis)
- Azure Computer Vision (for image analysis)
- Azure OpenAI (for text generation)
- Azure Document Intelligence (for document analysis)

## Setup

1. Clone this repository
2. Install the required packages:
   ```
   pip install python-dotenv requests
   ```
3. Create a `.env` file based on the `.env.template` file:
   ```
   cp .env.template .env
   ```
4. Update the `.env` file with your Azure AI service credentials

## Usage

The `home.py` script provides a command-line interface to interact with Azure AI services:

```bash
# Analyze text sentiment
python home.py --service text --input "I love Azure AI services, they are amazing!" --output results.json

# Analyze an image
python home.py --service image --input "https://example.com/image.jpg" --output results.json

# Generate text with Azure OpenAI
python home.py --service openai --input "Write a short poem about artificial intelligence" --output results.json

# Analyze a document
python home.py --service document --input "https://example.com/document.pdf" --output results.json
```

## Project Structure

- `home.py`: Main script for Azure AI integration
- `utilities/httpHelper.py`: Helper class for HTTP requests
- `.env.template`: Template for environment variables
