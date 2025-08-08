import requests
import msal
import os
import json
import pathspec

TENANT_ID = os.environ.get("INPUT_TENANT-ID")
CLIENT_ID = os.environ.get("INPUT_CLIENT-ID")
CLIENT_SECRET = os.environ.get("INPUT_CLIENT-SECRET")

SITE_ID = os.environ.get("INPUT_SITE-ID")
DRIVE_ID = os.environ.get("INPUT_DRIVE-ID")
LOCAL_DIRECTORY_PATH = os.environ.get("INPUT_LOCAL-DIRECTORY")
SHAREPOINT_BASE_FOLDER = os.environ.get("INPUT_SHAREPOINT-FOLDER")

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

def get_gitignore_spec(directory_path):
    """
    Finds .gitignore in the given directory and returns a pathspec object.
    Returns None if .gitignore is not found.
    """
    gitignore_path = os.path.join(directory_path, '.gitignore')
    if os.path.exists(gitignore_path):
        with open(gitignore_path, 'r') as f:
            print("Found .gitignore, applying rules.")
            return pathspec.PathSpec.from_lines('gitwildmatch', f)
    print("No .gitignore file found in the root directory.")
    return None

def upload_file_to_sharepoint(access_token, site_id, drive_id, local_file_path, sharepoint_file_path):
    """
    Uploads a single file to SharePoint using a resumable upload session.
    """
    if not os.path.exists(local_file_path):
        print(f"Error: The file '{local_file_path}' was not found.")
        return

    # The API endpoint for creating an upload session handles folder creation automatically.
    upload_session_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{sharepoint_file_path}:/createUploadSession"
    
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    
    session_payload = { "item": { "@microsoft.graph.conflictBehavior": "replace" } }

    session_response = requests.post(upload_session_url, headers=headers, data=json.dumps(session_payload))

    if session_response.status_code != 200:
        print(f"Error creating upload session for {sharepoint_file_path}: {session_response.status_code}")
        print(session_response.json())
        return

    upload_url = session_response.json().get('uploadUrl')
    if not upload_url:
        print(f"Could not get upload URL for {sharepoint_file_path}.")
        return

    file_size = os.path.getsize(local_file_path)
    # Use smaller chunks for potentially many small files to show progress faster
    chunk_size = 1 * 1024 * 1024  # 1 MB chunks
    chunks_uploaded = 0

    with open(local_file_path, 'rb') as f:
        while True:
            chunk = f.read(chunk_size)
            if not chunk:
                break # End of file

            start_index = chunks_uploaded
            end_index = start_index + len(chunk) - 1
            
            chunk_headers = {
                'Content-Length': str(len(chunk)),
                'Content-Range': f'bytes {start_index}-{end_index}/{file_size}'
            }
            
            upload_response = requests.put(upload_url, headers=chunk_headers, data=chunk)
            
            if not (200 <= upload_response.status_code <= 204):
                print(f"Error uploading chunk for {sharepoint_file_path}: {upload_response.status_code}")
                print(upload_response.json())
                return
            
            chunks_uploaded = end_index + 1
    
    # The final response on successful chunked upload is a 201 or 200
    print(f"Successfully uploaded: {sharepoint_file_path}")


if __name__ == "__main__":
    if "YOUR_TENANT_ID" in TENANT_ID or not os.path.isdir(LOCAL_DIRECTORY_PATH):
        print("Please fill in your Azure AD and SharePoint details in the script.")
        if not os.path.isdir(LOCAL_DIRECTORY_PATH):
            print(f"Error: The specified local directory does not exist: '{LOCAL_DIRECTORY_PATH}'")
    else:
        token = get_access_token(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
        
        if token:
            print("\nStarting directory upload...")
            gitignore_spec = get_gitignore_spec(LOCAL_DIRECTORY_PATH)
            
            # Walk through the local directory
            for root, dirs, files in os.walk(LOCAL_DIRECTORY_PATH, topdown=True):
                # Ignore the .git directory if it exists
                if '.git' in dirs:
                    dirs.remove('.git')

                # Filter out ignored directories
                if gitignore_spec:
                    # Filter dirs in-place for os.walk to not traverse them
                    dirs[:] = [d for d in dirs if not gitignore_spec.match_file(os.path.join(os.path.relpath(root, LOCAL_DIRECTORY_PATH), d))]

                for filename in files:
                    local_path = os.path.join(root, filename)
                    relative_path = os.path.relpath(local_path, LOCAL_DIRECTORY_PATH)

                    # Check if the file should be ignored
                    if gitignore_spec and gitignore_spec.match_file(relative_path):
                        print(f"Ignoring: {relative_path}")
                        continue
                    
                    # Construct the destination path for SharePoint (always use forward slashes)
                    sharepoint_path = os.path.join(SHAREPOINT_BASE_FOLDER, relative_path).replace(os.path.sep, '/')

                    # Upload the file
                    upload_file_to_sharepoint(
                        access_token=token,
                        site_id=SITE_ID,
                        drive_id=DRIVE_ID,
                        local_file_path=local_path,
                        sharepoint_file_path=sharepoint_path
                    )
            print("\nDirectory upload process finished.")
