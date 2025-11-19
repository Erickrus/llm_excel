import glob
import importlib
import inspect
import os
import time
import base64

FOLDER = "llm_temp"
POLL_INTERVAL = 0.5  # seconds

# Dynamically create funcs dict from functions.py
def create_funcs_dict():
    funcs = {}
    try:
        functions_module = importlib.import_module("functions")
        # Reload in case script changes while running (optional)
        importlib.reload(functions_module) 
        
        for name, obj in inspect.getmembers(functions_module, inspect.isfunction):
            if name.startswith("func_"):
                key = name[5:].upper()  # e.g., func_crawl -> CRAWL
                funcs[key] = obj
    except Exception as e:
        print(f"Error loading functions: {e}")
    return funcs

funcs = create_funcs_dict()

def handle_request(file_path: str):
    # Read as standard text (Base64 is just ASCII)
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read().strip()
    
    if not content:
        return
    
    # content format is: TYPE,BASE64_STRING
    request_type = None
    input_val_b64 = None
    
    if "," in content:
        request_type, input_val_b64 = content.split(",", 1)
    else:
        # Fallback if something went wrong
        return

    # 1. DECODE Input from Base64 -> UTF-8 String
    input_val = ""
    try:
        if input_val_b64:
            input_val = base64.b64decode(input_val_b64).decode('utf-8')
    except Exception as e:
        response_text = f"[Error decoding input: {e}]"
        write_response(file_path, response_text)
        return

    # Execute Logic
    print(f"Processing {request_type}...") # Debug log
    print(f"Input: {input_val}")
    
    if request_type in funcs:
        try:
            response_text = str(funcs[request_type](input_val))
        except Exception as e:
            response_text = f"[Error executing {request_type}: {e}]"
    else:
        response_text = "[Error: Unknown request type]"
    
    print(f"Response: {response_text}")
    write_response(file_path, response_text.replace("\n", "").replace("\r", ""))

def write_response(request_file_path, response_text):
    response_file = request_file_path.replace("request_", "response_")
    
    # 2. ENCODE Result to Base64
    try:
        # Encode string to bytes (utf-8), then bytes to base64 bytes, then to ascii string
        response_b64 = base64.b64encode(response_text.encode('utf-8')).decode('utf-8')
    except Exception as e:
        # Fallback in case of encoding error
        err = f"Encoding error: {e}"
        response_b64 = base64.b64encode(err.encode('utf-8')).decode('utf-8')

    # Write back
    with open(response_file, 'w', encoding='utf-8') as f:
        f.write(response_b64)
    
    # Cleanup request
    try:
        os.remove(request_file_path)
    except:
        pass

def main():
    print("Agent started. Watching folder:", FOLDER)
    if not os.path.exists(FOLDER):
        os.makedirs(FOLDER)

    # Clean start
    for txt_file in glob.glob(os.path.join(FOLDER, "*.txt")):
        try: os.remove(txt_file)
        except: pass

    while True:
        # Refresh functions list occasionally or just keep static? 
        # Kept static here for performance, move inside loop if you change functions.py often.
        
        for file in os.listdir(FOLDER):
            if file.startswith("request_") and file.endswith(".txt"):
                full_path = os.path.join(FOLDER, file)
                try:
                    # Slight delay to ensure VBA finished writing the file
                    time.sleep(0.1) 
                    handle_request(full_path)
                except Exception as e:
                    print(f"Error processing {file}: {e}")
        time.sleep(POLL_INTERVAL)

if __name__ == "__main__":
    main()
