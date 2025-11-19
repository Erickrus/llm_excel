import os
import requests
from langchain_community.llms import Ollama
from dotenv import load_dotenv

load_dotenv()

# Initialize LLM
# Increase timeout if model is slow to load
llm = Ollama(model="gemma3n:latest", timeout=120) 

def func_llm(prompt: str) -> str:
    try:
        # Ollama invoke usually returns a string directly in newer versions,
        # or an object. We handle both.
        res = llm.invoke(prompt)
        
        # Check if res is an object with .content (like AIMessage) or just a string
        if hasattr(res, 'content'):
            return res.content
        return str(res)
    except Exception as e:
        return f"[LLM Error: {str(e)}]"

def func_crawl(url: str) -> str:
    try:
        if not url.startswith("http"):
            url = "https://" + url
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        
        # Simple cleanup
        text = response.text
        return text[:2000] # Increased limit slightly
    except Exception as e:
        return f"[CRAWL error: {str(e)}]"

def func_read(file_path: str) -> str:
    try:
        # Remove quotes if user added them in Excel
        file_path = file_path.replace('"', '').replace("'", "")
        
        if not os.path.exists(file_path):
            return "[Error: File not found]"
            
        with open(file_path, 'r', encoding='utf-8') as f:
            return f.read()[:2000]
    except Exception as e:
        return f"[READ error: {str(e)}]"

def func_agent(message: str) -> str:
    return f"Run agent successfully received: {message}"
