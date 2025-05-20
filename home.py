#!/usr/bin/env python
# filepath: c:\Users\kevin.li\OneDrive - GREEN DOT CORPORATION\Documents\GitHub\Azure\AzureAI\home.py

import os
import json
import argparse
from azureAIService import AzureAIService
from utilities.httpHelper import httpHelper
from dotenv import load_dotenv
import sys

# Load environment variables and verify .env file
dotenv_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.env')
if os.path.exists(dotenv_path):
    print(f"Found .env file at: {dotenv_path}")
    load_dotenv(dotenv_path)
else:
    print(f"WARNING: .env file not found at: {dotenv_path}")
    print("Creating .env file from template...")
    try:
        template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.env.template')
        if os.path.exists(template_path):
            with open(template_path, 'r') as template_file:
                template_content = template_file.read()
            with open(dotenv_path, 'w') as env_file:
                env_file.write(template_content)
            print(f"Created .env file at: {dotenv_path}")
            print("Please edit the .env file to add your Azure AI service credentials.")
            load_dotenv(dotenv_path)
        else:
            print(f"ERROR: .env.template file not found at: {template_path}")
    except Exception as e:
        print(f"ERROR creating .env file: {str(e)}")
def main():
    """
    Main function to parse arguments and execute Azure AI operations.
    """
    # Check if any arguments were provided
    import sys
    if len(sys.argv) == 1:
        # Interactive mode
        print("Azure AI Services Demo - Interactive Mode")
        print("----------------------------------------")
        print("Select a service:")
        print("1. Text Analysis")
        print("2. Image Analysis")
        print("3. OpenAI Text Generation")
        print("4. Document Analysis")
        
        choice = input("Enter your choice (1-4): ")
        service_map = {
            "1": "text",
            "2": "image",
            "3": "openai",
            "4": "document"
        }
        
        if choice in service_map:
            service = service_map[choice]
            input_text = input(f"Enter {service} input: ")
            output_path = input("Enter output file path (press Enter to skip): ")
            
            args = type('Args', (), {
                'service': service,
                'input': input_text,
                'output': output_path if output_path else None
            })
            
            print(f"\nSelected service: {args.service}")
        else:
            print("Invalid choice. Exiting.")
            return
    else:
        # Command-line mode
        parser = argparse.ArgumentParser(description="Azure AI Services Demo")
        parser.add_argument("--service", choices=["text", "image", "openai", "document"], 
                          help="Select Azure AI service to use", required=True)
        parser.add_argument("--input", help="Input text, URL, or prompt", required=True)
        parser.add_argument("--output", help="Output file path (optional)")
        
        args = parser.parse_args()
        print(f"Selected service: {args.service}")
    
    # Initialize Azure AI service
    azure_ai = AzureAIService()

    # Process based on selected service
    result = None
    if args.service == "text":
        result = azure_ai.analyze_text(args.input)
    elif args.service == "image":
        result = azure_ai.analyze_image(args.input)
    elif args.service == "openai":
        result = azure_ai.generate_text(args.input)
    elif args.service == "document":
        result = azure_ai.analyze_document(args.input)
    
    # Print result
    if result:
        print("Result:")
        print(json.dumps(result, indent=2))
        
        # Save to file if output path is provided
        if args.output:
            with open(args.output, 'w') as f:
                json.dump(result, f, indent=2)
            print(f"Results saved to {args.output}")

if __name__ == "__main__":
    main()