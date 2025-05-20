
import json
import os
from utilities.httpHelper import httpHelper


class AzureAIService:
    """
    A class to interact with various Azure AI services.
    """
    def __init__(self):
        # Load configurations from environment variables with validation
        self.language_endpoint = os.environ.get("AZURE_LANGUAGE_ENDPOINT")
        self.language_key = os.environ.get("AZURE_LANGUAGE_KEY")
        self.vision_endpoint = os.environ.get("AZURE_VISION_ENDPOINT")
        self.vision_key = os.environ.get("AZURE_VISION_KEY")
        self.openai_endpoint = os.environ.get("AZURE_OPENAI_ENDPOINT")
        self.openai_key = os.environ.get("AZURE_OPENAI_KEY")
        self.openai_deployment = os.environ.get("AZURE_OPENAI_DEPLOYMENT", "gpt-35-turbo")
        self.document_endpoint = os.environ.get("AZURE_DOCUMENT_ENDPOINT")
        self.document_key = os.environ.get("AZURE_DOCUMENT_KEY")
        
        # Print debug info about environment variables
        print("\nEnvironment Variables Status:")
        env_vars = [
            ("AZURE_LANGUAGE_ENDPOINT", self.language_endpoint),
            ("AZURE_LANGUAGE_KEY", self.language_key and "***" or None),
            ("AZURE_VISION_ENDPOINT", self.vision_endpoint),
            ("AZURE_VISION_KEY", self.vision_key and "***" or None),
            ("AZURE_OPENAI_ENDPOINT", self.openai_endpoint),
            ("AZURE_OPENAI_KEY", self.openai_key and "***" or None),
            ("AZURE_OPENAI_DEPLOYMENT", self.openai_deployment),
            ("AZURE_DOCUMENT_ENDPOINT", self.document_endpoint),
            ("AZURE_DOCUMENT_KEY", self.document_key and "***" or None)
        ]
        
        for name, value in env_vars:
            status = "✓ Set" if value else "✗ Not set"
            print(f"{name}: {status}")
        print("")
    
    def analyze_text(self, text):
        """
        Analyzes text using Azure Language Service for sentiment and key phrases.
        
        Args:
            text (str): The text to analyze
    
        Returns:
            dict: The analysis results
        """
        print(f"Analyzing text: {text[:50]}...")
        
        url = f"{self.language_endpoint}/text/analytics/v3.1/sentiment"
        headers = {
            "Ocp-Apim-Subscription-Key": self.language_key,
            "Content-Type": "application/json"
        }
        
        data = {
            "documents": [
                {
                    "id": "1",
                    "language": "en",
                    "text": text
                }
            ]
        }
        
        try:
            response = httpHelper.post(url, data)
            return json.loads(response)
        except Exception as e:
            print(f"Error analyzing text: {str(e)}")
            return {"error": str(e)}
    
    def analyze_image(self, image_url):
        """
        Analyzes an image using Azure Computer Vision.
        
        Args:
            image_url (str): URL of the image to analyze
            
        Returns:
            dict: Image analysis results
        """
        print(f"Analyzing image at: {image_url}")
        
        url = f"{self.vision_endpoint}/computervision/imageanalysis:analyze?api-version=2023-04-01-preview&features=caption,read,denseCaptions,objects,people,smartCrops,tags"
        headers = {
            "Ocp-Apim-Subscription-Key": self.vision_key,
            "Content-Type": "application/json"
        }
        
        data = {"url": image_url}
        
        try:
            response = httpHelper.post(url, data)
            return json.loads(response)
        except Exception as e:
            print(f"Error analyzing image: {str(e)}")
            return {"error": str(e)}
    
    def generate_text(self, prompt, max_tokens=1000):
        """
        Generates text using Azure OpenAI.
        
        Args:
            prompt (str): The text prompt
            max_tokens (int): Maximum tokens to generate
            
        Returns:
            dict: Generated text response
        """
        print(f"Generating text for prompt: {prompt[:50]}...")
        
        url = f"{self.openai_endpoint}"
        headers = {
            "api-key": self.openai_key,
            "Content-Type": "application/json",
            "Location":"eastus",
            "deployment": self.openai_deployment
        }
        
        data = {
            "prompt": prompt,
            "max_tokens": max_tokens,
            "temperature": 0.7,
            "top_p": 0.95
        }
        
        try:
            response = httpHelper.post(url, data)
            return json.loads(response)
        except Exception as e:
            print(f"Error generating text: {str(e)}")
            return {"error": str(e)}
    
    def analyze_document(self, document_url):
        """
        Analyzes a document using Azure Document Intelligence.
        
        Args:
            document_url (str): URL of the document to analyze
            
        Returns:
            dict: Document analysis results
        """
        print(f"Analyzing document at: {document_url}")
        
        url = f"{self.document_endpoint}/documentintelligence/documentModels/prebuilt-layout:analyze?api-version=2023-07-31"
        headers = {
            "Ocp-Apim-Subscription-Key": self.document_key,
            "Content-Type": "application/json"
        }
        
        data = {"urlSource": document_url}
        
        try:
            response = httpHelper.post(url, data)
            return json.loads(response)
        except Exception as e:
            print(f"Error analyzing document: {str(e)}")
            return {"error": str(e)}