import re
import os
import io
import mimetypes
import time
import requests
from base64 import b64encode
from datetime import datetime
from googleapiclient.http import MediaIoBaseDownload
from googleapiclient.errors import HttpError

from constants import (
    drive_service, WP_URL, WP_USER, WP_PASSWORD,
    GREEN, YELLOW, RED, BLUE, BOLD, ENDC
)

def download_image(file_id):
    """Download image from Google Drive."""
    try:
        request = drive_service.files().get_media(fileId=file_id)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        return fh.getvalue()
    except HttpError as error:
        print(f"Image download failed: {error}")
        return None
    
def process_image_from_url(image_url, caption, doc_id):
    """Process image from a Google Drive URL."""
    if not image_url:
        return None
        
    # Define WordPress supported image formats
    SUPPORTED_FORMATS = {
        'image/png': '.png',
        'image/jpeg': '.jpg',
        'image/jpg': '.jpg',
        'image/gif': '.gif',
        'image/webp': '.webp',
        'image/heic': '.heic',
        'image/heif': '.heif'
    }
        
    file_id = extract_file_id(image_url)
    if not file_id:
        print(f"Failed to extract file ID from URL: {image_url}")
        return None
        
    print(f"Downloading image with file ID: {file_id}")
    
    # Get file metadata first to determine the file type
    try:
        file_metadata = drive_service.files().get(fileId=file_id, fields="name,mimeType").execute()
        file_name = file_metadata.get('name', f"image_{doc_id}")
        file_mime_type = file_metadata.get('mimeType', '')
        
        print(f"File name from Drive: {file_name}")
        print(f"File MIME type from Drive: {file_mime_type}")
        
        # Determine file extension from mime type or original filename
        file_ext = None
        
        # First try to get extension from the file name
        if '.' in file_name:
            original_ext = os.path.splitext(file_name)[1].lower()  # Get extension including the dot
            # Check if this is a supported extension
            for mime, ext in SUPPORTED_FORMATS.items():
                if original_ext == ext or original_ext == '.jpeg' and ext == '.jpg':
                    file_ext = original_ext
                    break
        
        # If extension not determined from filename, try from mime type
        if not file_ext and file_mime_type in SUPPORTED_FORMATS:
            file_ext = SUPPORTED_FORMATS[file_mime_type]
            print(f"Using extension {file_ext} based on MIME type")
        
        # If we still don't have a supported extension, we need to use fallback options
        if not file_ext:
            print(f"Unsupported image format detected: {file_mime_type}")
            print(f"Original filename: {file_name}")
            print("WordPress only supports: PNG, JPG/JPEG, GIF, WebP, HEIC, and HEIF")
            return None  # This will trigger the fallback in the main function
            
        print(f"Using file extension: {file_ext}")
        
    except Exception as e:
        print(f"Warning: Could not determine image format: {str(e)}")
        print("Cannot verify if image format is supported. Trying fallback options...")
        return None  # Trigger fallback
    
    # Download the file
    image_data = download_image(file_id)
    if not image_data:
        print(f"Failed to download image data from file ID: {file_id}")
        return None
        
    filename = f"featured_image_{doc_id}{file_ext}"
    print(f"Uploading image: {filename}")
    
    # Use improved upload function with retries
    return upload_image_to_wordpress(
        image_data,
        caption,
        filename,
        max_retries=2,
        retry_delay=2
    )

def handle_image_fallback(caption, doc_id):
    """Handle image upload fallback when the initial upload fails."""
    # Define WordPress supported image formats
    SUPPORTED_FORMATS = ['.png', '.jpg', '.jpeg', '.gif', '.webp', '.heic', '.heif']
    
    print(f"\n{YELLOW}{BOLD}Image upload fallback options:{ENDC}")
    print("1. Enter a new Google Drive URL")
    print("2. Provide a local file path")
    print("3. Skip image upload (continue without image)")
    
    choice = input(f"\n{BLUE}Select an option (1-3): {ENDC}").strip()
    
    if choice == '1':
        # Option 1: New Google Drive URL
        new_url = input(f"{BLUE}Enter new Google Drive URL: {ENDC}").strip()
        if new_url:
            return process_image_from_url(new_url, caption, doc_id)
        else:
            print(f"{RED}No URL provided. Skipping image upload.{ENDC}")
            return None
            
    elif choice == '2':
        # Option 2: Local file path
        local_path = input(f"{BLUE}Enter local file path: {ENDC}").strip()
        if os.path.exists(local_path):
            try:
                # Check file extension first
                file_ext = os.path.splitext(local_path)[1].lower()
                if file_ext not in SUPPORTED_FORMATS:
                    print(f"{RED}Unsupported file format: {file_ext}{ENDC}")
                    print(f"{YELLOW}WordPress only supports: {', '.join(SUPPORTED_FORMATS)}{ENDC}")
                    print(f"{YELLOW}Please select a different file.{ENDC}")
                    # Recursive call to try again
                    return handle_image_fallback(caption, doc_id)
                
                # If we get here, the file format is supported
                with open(local_path, 'rb') as file:
                    image_data = file.read()
                
                # Preserve the original filename with extension
                original_filename = os.path.basename(local_path)
                filename = f"featured_image_{doc_id}_{original_filename}"
                print(f"Uploading local image: {filename}")
                
                # Get mime type from the file extension
                mime_type = mimetypes.guess_type(local_path)[0] or 'image/jpeg'
                print(f"Detected mime type: {mime_type}")
                
                return upload_image_to_wordpress(
                    image_data,
                    caption,
                    filename,
                    max_retries=3,
                    retry_delay=3
                )
            except Exception as e:
                print(f"{RED}Error reading local file: {str(e)}{ENDC}")
                # Recursive call to try again
                return handle_image_fallback(caption, doc_id)
        else:
            print(f"{RED}File not found: {local_path}{ENDC}")
            # Recursive call to try again
            return handle_image_fallback(caption, doc_id)
    
    else:
        # Option 3 or invalid input: Skip image upload
        print(f"{YELLOW}Skipping image upload.{ENDC}")
        return None

def extract_file_id(url):
    """Extract Google Drive file ID from URL."""
    # Support both file/d/ and open?id= formats
    file_patterns = [
        r'/file/d/([a-zA-Z0-9_-]+)',
        r'id=([a-zA-Z0-9_-]+)',
        r'/open\?id=([a-zA-Z0-9_-]+)'
    ]

    for pattern in file_patterns:
        match = re.search(pattern, url)
        if match:
            return match.group(1)
    return None

def upload_image_to_wordpress(image_data, caption, filename, max_retries=3, retry_delay=3):
    """Upload image to WordPress media library with retry logic and improved error handling."""
    if not image_data:
        print("No image data provided")
        return None

    mime_type = mimetypes.guess_type(filename)[0] or 'image/jpeg'

    for attempt in range(1, max_retries + 1):
        try:
            # Ensure unique filename
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"{timestamp}_{filename}"

            headers = {
                'Authorization': f'Basic {b64encode(f"{WP_USER}:{WP_PASSWORD}".encode()).decode()}'
            }

            files = {
                'file': (filename, image_data, mime_type),
            }

            data = {
                'title': filename,
                'caption': caption,
                'alt_text': caption
            }

            print(f"Attempt {attempt}/{max_retries}: Uploading image '{filename}'")
            
            response = requests.post(
                f'{WP_URL}/wp/v2/media',
                headers=headers,
                files=files,
                data=data,
                timeout=30
            )

            if response.status_code == 201:
                media_data = response.json()
                print(f"Successfully uploaded image: {media_data.get('source_url')}")
                return media_data.get('id')
            else:
                print(f"Image upload failed with status {response.status_code}: {response.text}")
                if attempt < max_retries:
                    print(f"Waiting {retry_delay} seconds before retrying...")
                    time.sleep(retry_delay)
                else:
                    print(f"Maximum retry attempts reached. Image upload failed.")
                    return None

        except Exception as e:
            print(f"Image upload failed: {str(e)}")
            if attempt < max_retries:
                print(f"Waiting {retry_delay} seconds before retrying...")
                time.sleep(retry_delay)
            else:
                print(f"Maximum retry attempts reached. Image upload failed.")
                return None
    
    return None
