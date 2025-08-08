import requests
import msal
import os
import json
import pathspec
import time

TENANT_ID = os.environ.get("TENANT_ID")
CLIENT_ID = os.environ.get("CLIENT_ID")
CLIENT_SECRET = os.environ.get("CLIENT_SECRET")

SITE_ID = os.environ.get("SITE_ID")
DRIVE_ID = os.environ.get("DRIVE_ID")
LOCAL_DIRECTORY_PATH = os.environ.get("LOCAL_DIRECTORY_PATH")
SHAREPOINT_BASE_FOLDER = os.environ.get("SHAREPOINT_BASE_FOLDER")
SYNC_DELETIONS = os.environ.get("SYNC_DELETIONS", 'false').lower() == 'true'

# --- Script Logic ---

def get_access_token(tenant_id, client_id, client_secret):
    """
    Authenticates with Azure AD using client credentials flow and returns an access token.
    MSAL handles token caching automatically.
    """
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = msal.ConfidentialClientApplication(
        client_id=client_id,
        authority=authority,
        client_credential=client_secret
    )
    
    scopes = ["https://graph.microsoft.com/.default"]
    result = app.acquire_token_for_client(scopes=scopes)
        
    if "access_token" in result:
        print("Access token acquired successfully.")
        return result['access_token']
    else:
        print("Error acquiring access token.")
        print(result.get("error"))
        print(result.get("error_description"))
        print(result.get("correlation_id"))
        return None

def get_remote_files_recursive(access_token, site_id, drive_id, item_path=""):
    """
    Recursively lists all files in a SharePoint directory and returns a dictionary
    mapping their relative path to their item ID.
    """
    headers = {'Authorization': f'Bearer {access_token}'}
    # Properly format the root or subfolder path for the API call
    if not item_path:
        # If we are at the root of the upload folder
        api_path = f"root:/{SHAREPOINT_BASE_FOLDER}"
    else:
        # If we are in a subfolder
        api_path = f"root:/{SHAREPOINT_BASE_FOLDER}/{item_path}"
    
    # URL encode the path to handle special characters
    encoded_api_path = requests.utils.quote(api_path)
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/{encoded_api_path}:/children"
    
    remote_files = {}
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status() # Raise an exception for bad status codes
        items = response.json().get('value', [])
        
        for item in items:
            current_path = os.path.join(item_path, item['name']).replace(os.path.sep, '/')
            if 'file' in item:
                remote_files[current_path] = item['id']
            elif 'folder' in item:
                # Give Graph API a moment before the next recursive call
                time.sleep(0.1) 
                remote_files.update(get_remote_files_recursive(access_token, site_id, drive_id, current_path))
                
    except requests.exceptions.HTTPError as e:
        # If the base folder doesn't exist, it's not an error; it just means there are no files to list.
        if e.response.status_code == 404:
            print(f"SharePoint folder '{SHAREPOINT_BASE_FOLDER}' not found. No remote files to compare.")
            return {}
        else:
            print(f"Error listing remote files at '{item_path}': {e}")
            print(f"Response: {e.response.text}")
            exit(1) # Exit on other HTTP errors
            
    return remote_files

def delete_sharepoint_item(access_token, site_id, drive_id, item_id):
    """Deletes an item (file or folder) from SharePoint by its ID."""
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/items/{item_id}"
    headers = {'Authorization': f'Bearer {access_token}'}
    response = requests.delete(url, headers=headers)
    if response.status_code == 204:
        print(f"Successfully deleted item {item_id}")
    else:
        print(f"Error deleting item {item_id}: {response.status_code} - {response.text}")

def upload_file_to_sharepoint(access_token, site_id, drive_id, local_file_path, sharepoint_file_path):
    """Uploads a single file to SharePoint using a resumable upload session."""
    # ... (This function remains the same as before)
    upload_session_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{sharepoint_file_path}:/createUploadSession"
    headers = {'Authorization': f'Bearer {access_token}', 'Content-Type': 'application/json'}
    session_payload = {"item": {"@microsoft.graph.conflictBehavior": "replace"}}
    session_response = requests.post(upload_session_url, headers=headers, data=json.dumps(session_payload))
    if session_response.status_code != 200:
        print(f"Error creating upload session for {sharepoint_file_path}: {session_response.status_code} - {session_response.json()}")
        return
    upload_url = session_response.json().get('uploadUrl')
    file_size = os.path.getsize(local_file_path)
    chunk_size = 4 * 1024 * 1024
    with open(local_file_path, 'rb') as f:
        start_index = 0
        while True:
            chunk = f.read(chunk_size)
            if not chunk: break
            end_index = start_index + len(chunk) - 1
            chunk_headers = {'Content-Length': str(len(chunk)), 'Content-Range': f'bytes {start_index}-{end_index}/{file_size}'}
            upload_response = requests.put(upload_url, headers=chunk_headers, data=chunk)
            if not (200 <= upload_response.status_code <= 204):
                print(f"Error uploading chunk for {sharepoint_file_path}: {upload_response.status_code} - {upload_response.json()}")
                return
            start_index = end_index + 1
    print(f"Successfully uploaded: {sharepoint_file_path}")

# --- Main Execution Block ---
if __name__ == "__main__":
    # --- 1. Get Local File List ---
    print("Gathering local file list...")
    gitignore_path = os.path.join(LOCAL_DIRECTORY_PATH, '.gitignore')
    spec = None
    if os.path.exists(gitignore_path):
        with open(gitignore_path, 'r') as f:
            spec = pathspec.PathSpec.from_lines('gitwildmatch', f)
            
    local_files = set()
    for root, dirs, files in os.walk(LOCAL_DIRECTORY_PATH, topdown=True):
        if '.git' in dirs: dirs.remove('.git')
        if spec:
            dirs[:] = [d for d in dirs if not spec.match_file(os.path.join(os.path.relpath(root, LOCAL_DIRECTORY_PATH), d))]
        for filename in files:
            local_path = os.path.join(root, filename)
            relative_path = os.path.relpath(local_path, LOCAL_DIRECTORY_PATH)
            if not spec or not spec.match_file(relative_path):
                local_files.add(relative_path.replace(os.path.sep, '/'))

    # --- 2. Authenticate and Handle Deletions (if enabled) ---
    print("Authenticating with Microsoft Graph...")
    token = get_access_token(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
    
    if SYNC_DELETIONS:
        print("\nSync deletions enabled. Comparing remote files with local files...")
        remote_files_map = get_remote_files_recursive(token, SITE_ID, DRIVE_ID)
        remote_files_set = set(remote_files_map.keys())
        
        files_to_delete = remote_files_set - local_files
        
        if files_to_delete:
            print(f"\nFound {len(files_to_delete)} files to delete from SharePoint:")
            for file_path in sorted(list(files_to_delete)):
                print(f" - Deleting: {file_path}")
                item_id = remote_files_map[file_path]
                delete_sharepoint_item(token, SITE_ID, DRIVE_ID, item_id)
        else:
            print("SharePoint directory is already in sync. No files to delete.")
    else:
        print("\nSync deletions is disabled. Skipping deletion step.")

    # --- 3. Upload Files ---
    print("\nStarting file uploads...")
    for relative_path in sorted(list(local_files)):
        local_path = os.path.join(LOCAL_DIRECTORY_PATH, relative_path)
        sharepoint_path = os.path.join(SHAREPOINT_BASE_FOLDER, relative_path).replace(os.path.sep, '/')
        upload_file_to_sharepoint(token, SITE_ID, DRIVE_ID, local_path, sharepoint_path)
        
    print("\nProcess finished.")