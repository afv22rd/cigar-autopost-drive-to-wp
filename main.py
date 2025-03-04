import re
import requests
import os
import io
import time
from googleapiclient.discovery import build
from google.oauth2 import service_account
from googleapiclient.http import MediaIoBaseDownload
from googleapiclient.errors import HttpError
import mimetypes
from datetime import datetime
from base64 import b64encode
from dotenv import load_dotenv

load_dotenv()

# Configuration
GOOGLE_CREDENTIALS_FILE = os.getenv('GOOGLE_CREDENTIALS_FILE')
WP_URL = os.getenv('WP_URL')
WP_USER = os.getenv('WP_USER')
WP_PASSWORD = os.getenv('WP_PASSWORD')

# Google Sheets green color (normalized to 0-1 range)
GREEN_COLOR = {'red': 0.5764706, 'green': 0.76862746, 'blue': 0.49019608}

# Google API setup
SCOPES = ['https://www.googleapis.com/auth/spreadsheets',
          'https://www.googleapis.com/auth/drive',
          'https://www.googleapis.com/auth/documents.readonly']
creds = service_account.Credentials.from_service_account_file(
    GOOGLE_CREDENTIALS_FILE, scopes=SCOPES)
sheets_service = build('sheets', 'v4', credentials=creds)
docs_service = build('docs', 'v1', credentials=creds)
drive_service = build('drive', 'v3', credentials=creds)

def get_eligible_rows(sheet_id):
    """
    Retrieve rows from Google Sheet where:
    - 'Ready To Post' (Column B) is checked
    - 'Online' (Column D) is NOT checked
    - Track sections from Column A
    """
    sheet = sheets_service.spreadsheets().get(
        spreadsheetId=sheet_id,
        includeGridData=True,
        fields="sheets(data(rowData(values(formattedValue,effectiveFormat,hyperlink,textFormatRuns,userEnteredFormat.backgroundColor))))"
    ).execute()
    rows = sheet['sheets'][0]['data'][0]['rowData']

    eligible_rows = []
    current_section = "Uncategorized"  # Default section
    
    for row_idx, row in enumerate(rows[7:]):  # Skip header rows
        try:
            actual_row_num = row_idx + 8  # Actual row number in spreadsheet
            print(f"\nAnalyzing Row {actual_row_num}:")
            
            # Skip empty rows
            if not row.get('values'):
                print(f"Row {actual_row_num}: Empty row - skipping")
                continue
            
            values = row.get('values', [])
            
            # Check if this is a section header in Column A
            if len(values) > 0 and 'formattedValue' in values[0]:
                section_text = values[0].get('formattedValue', '').strip()
                if section_text and not any('formattedValue' in values[i] for i in [1, 3, 4] if i < len(values)):
                    # This looks like a section header - no data in columns B, D, E
                    current_section = section_text
                    print(f"Row {actual_row_num}: Found section header: {current_section}")
                    continue  # Skip processing this row as a content item
            
            # Debug print for first few columns
            print(f"Section: {current_section}")
            print(f"Column B (Ready): {values[1].get('formattedValue', 'Empty') if len(values) > 1 else 'Missing'}")
            print(f"Column D (Online): {values[3].get('formattedValue', 'Empty') if len(values) > 3 else 'Missing'}")
            print(f"Column E (Story): {values[4].get('formattedValue', 'Empty') if len(values) > 4 else 'Missing'}")

            # Check Ready to Post status (Column B)
            ready_cell = values[1].get('formattedValue', '').upper() if len(values) > 1 else ''
            is_ready = ready_cell in ['TRUE', '✓', 'YES', '1']
            if not is_ready:
                print(f"Row {actual_row_num}: Not ready to post ({ready_cell}) - skipping")
                continue

            # Check Online status (Column D)
            online_cell = values[3].get('formattedValue', '').upper() if len(values) > 3 else ''
            is_online = online_cell in ['TRUE', '✓', 'YES', '1']
            if is_online:
                print(f"Row {actual_row_num}: Already online - skipping")
                continue

            # Get Story URL (Column E)
            story_cell = values[4] if len(values) > 4 else {}
            story_url = None

            print(f"Row {actual_row_num}: Analyzing story cell (Column E):")
            print(f"  Cell content: {story_cell}")

            # Method 1: Try to get URL from textFormatRuns
            if 'textFormatRuns' in story_cell:
                print(f"  textFormatRuns found")
                for run in story_cell['textFormatRuns']:
                    if 'format' in run and 'link' in run['format']:
                        story_url = run['format']['link']['uri']
                        print(f"  Found URL from textFormatRuns: {story_url}")
                        break

            # Method 2: Try to get URL from hyperlink property
            if not story_url and 'hyperlink' in story_cell:
                story_url = story_cell['hyperlink']
                print(f"  Found URL from hyperlink property: {story_url}")

            # Method 3: Look for URL patterns in text
            if not story_url and 'formattedValue' in story_cell:
                url_match = re.search(r'https?://[^\s]+', story_cell['formattedValue'])
                if url_match:
                    story_url = url_match.group()
                    print(f"  Found URL from text pattern: {story_url}")

            if not story_url:
                print(f"  No valid story URL found - skipping")
                continue

            print(f"  Using URL: {story_url}")

            # Get Author (Column H)
            author_cell = values[7] if len(values) > 7 else None
            author_names = []
            if author_cell and 'formattedValue' in author_cell:
                author_name = author_cell.get('formattedValue', '').strip()
                if author_name:
                    author_names = [name.strip() for name in author_name.split(',')]
                    print(f"Row {actual_row_num}: Found author(s): {', '.join(author_names)}")

            # Get Categories (Column O)
            categories = []
            if len(values) > 14:  # Column O exists
                categories_cell = values[14]
                if categories_cell and 'formattedValue' in categories_cell:
                    categories_text = categories_cell.get('formattedValue', '').strip()
                    if categories_text:
                        categories = [cat.strip() for cat in categories_text.split(',')]
                        print(f"Row {actual_row_num}: Found categories: {', '.join(categories)}")

            # If no categories found (either column missing or empty), use section
            if not categories:
                print(f"Row {actual_row_num}: No categories found. Setting to default section category: {current_section}")
                categories = [current_section]
                
            # Get Photographer info (Column N)
            photographer_name = None
            if len(values) > 13:  # Column N is index 13
                photographer_cell = values[13]
                if photographer_cell and 'formattedValue' in photographer_cell:
                    photographer_name = photographer_cell.get('formattedValue', '').strip()
                    if photographer_name:
                        print(f"Row {actual_row_num}: Found photographer: {photographer_name}")

            # Add to eligible rows
            print(f"Row {actual_row_num}: Adding to eligible rows (Section: {current_section})")
            eligible_rows.append({
                'row': actual_row_num,
                'doc_url': story_url,
                'author_names': author_names,
                'categories': categories,
                'photographer_name': photographer_name,
                'online_cell': f"D{actual_row_num}",
                'section': current_section  # Store the section with each row
            })

        except Exception as e:
            print(f"Error processing row {row_idx + 8}: {str(e)}")
            continue

    return eligible_rows

def parse_google_doc(doc_id):
    """Extract sections from Google Doc with improved parsing."""
    doc = docs_service.documents().get(documentId=doc_id).execute()
    content = doc['body']['content']

    sections = {
        'Headline': '',
        'Featured image': '',
        'Cutlines': '',
        'Redaction': '',
    }

    current_section = None

    for element in content:
        if 'paragraph' in element:
            elements = element['paragraph']['elements']
            text = ''.join([e.get('textRun', {}).get('content', '') for e in elements]).strip()

            # Check for section headers
            for section in sections:
                header_pattern = f"{section}:"
                if text.startswith(header_pattern):
                    current_section = section
                    sections[current_section] = text[len(header_pattern):].strip()
                    break
            else:
                if current_section and text:
                    sections[current_section] += "\n" + text

    return sections

def get_or_create_author_id(author_name):
    """
    Search WordPress users by name and return their user ID.
    If not found, create a new user with the provided name.
    """
    try:
        # Clean up author name (remove any extra whitespace or newlines)
        author_name = author_name.strip()
        
        # Split authors by comma and clean whitespace (for the first one)
        primary_author = author_name.split(',')[0].strip()
        
        # Log if multiple authors detected
        if ',' in author_name:
            co_authors = [name.strip() for name in author_name.split(',')[1:]]
            print(f"WARNING: Multiple authors detected: {author_name}")
            print(f"Using primary author '{primary_author}'. Please manually add these co-authors: {', '.join(co_authors)}")

        headers = {
            'Authorization': f'Basic {b64encode(f"{WP_USER}:{WP_PASSWORD}".encode()).decode()}'
        }

        # Use the specific users endpoint
        users_endpoint = f'{WP_URL}/wp/v2/users'

        # Search for the primary author
        params = {'search': primary_author}
        response = requests.get(
            users_endpoint,
            headers=headers,
            params=params,
            timeout=10
        )

        if response.status_code == 200:
            users = response.json()
            if users:
                # Look for exact match first (case-insensitive)
                for user in users:
                    if user['name'].lower() == primary_author.lower():
                        print(f"Found exact match for primary author '{primary_author}' with ID: {user['id']}")
                        return user['id']

                # If no exact match, return first result
                print(f"Found partial match for primary author '{primary_author}' with ID: {users[0]['id']}")
                return users[0]['id']
                
        # If we get here, the author was not found
        print(f"Author '{primary_author}' not found. Creating new user...")
        return create_wordpress_user(primary_author)

    except Exception as e:
        print(f"Error searching for author: {e}")
        return None

def create_wordpress_user(full_name):
    """Create a new WordPress user for an author."""
    try:
        # Parse name
        name_parts = full_name.strip().split()
        if len(name_parts) < 2:
            print(f"ERROR: Author name '{full_name}' doesn't have both first and last name")
            return None
            
        first_name = name_parts[0]
        last_name = ' '.join(name_parts[1:])  # Join remaining parts as last name
        
        # Create username: name.lastname (lowercase, no spaces)
        username = f"{first_name.lower()}.{last_name.lower().replace(' ', '')}"
        
        # Create email using username (alternate between domains if needed)
        import random
        import string
        domain = random.choice(["nogood.com", "nogood.net"])
        email = f"{username}@{domain}"
        
        # Generate random password
        special_chars = "!\"#$%&'()*+,-./:;<=>?@[]^_{}|~"
        password = ''.join(random.choices(string.ascii_letters + string.digits + special_chars, k=20))
        
        # Prepare user data
        user_data = {
            'username': username,
            'email': email,
            'first_name': first_name,
            'last_name': last_name,
            'roles': ['staff-writer'],
            'password': password
        }
        
        # Send request to WordPress API
        headers = {
            'Content-Type': 'application/json',
            'Authorization': f'Basic {b64encode(f"{WP_USER}:{WP_PASSWORD}".encode()).decode()}'
        }
        
        response = requests.post(
            f'{WP_URL}/wp/v2/users',
            json=user_data,
            headers=headers,
            timeout=15
        )
        
        if response.status_code == 201:
            new_user = response.json()
            print(f"Successfully created new user '{full_name}':")
            print(f"  Username: {username}")
            print(f"  Email: {email}")
            print(f"  Password: {password}")
            print(f"  User ID: {new_user['id']}")
            return new_user['id']
        else:
            print(f"Failed to create user for '{full_name}'. Response: {response.status_code} - {response.text}")
            return None
            
    except Exception as e:
        print(f"Error creating WordPress user: {e}")
        return None

def get_category_ids(categories_list):
    """Search WordPress categories by name and return their IDs."""
    try:
        headers = {
            'Authorization': f'Basic {b64encode(f"{WP_USER}:{WP_PASSWORD}".encode()).decode()}'
        }

        category_ids = []

        # Use the categories endpoint
        categories_endpoint = f'{WP_URL}/wp/v2/categories'
        
        # Get all categories first (to avoid multiple API calls)
        response = requests.get(
            categories_endpoint,
            headers=headers,
            params={'per_page': 100},  # Increase if you have more categories
            timeout=10
        )
        
        if response.status_code != 200:
            print(f"Failed to fetch categories. Response: {response.text}")
            return []
            
        all_categories = response.json()
        print(f"Found {len(all_categories)} total categories in WordPress")

        # Process each requested category
        for name in categories_list:
            original_name = name.strip()
            found = False
            
            # Try direct match first (case insensitive)
            for category in all_categories:
                if category['name'].lower() == original_name.lower():
                    category_ids.append(category['id'])
                    print(f"Found exact match for category '{original_name}' → '{category['name']}' with ID: {category['id']}")
                    found = True
                    break
            
            if found:
                continue
                
            # Try standardized versions (replace '&' with 'and' and vice versa)
            standardized_with_and = original_name.replace('&', 'and')
            standardized_with_ampersand = original_name.replace(' and ', ' & ')
            
            for category in all_categories:
                cat_name_lower = category['name'].lower()
                if (cat_name_lower == standardized_with_and.lower() or 
                    cat_name_lower == standardized_with_ampersand.lower()):
                    category_ids.append(category['id'])
                    print(f"Found match using standardized name for '{original_name}' → '{category['name']}' with ID: {category['id']}")
                    found = True
                    break
            
            if found:
                continue
                
            # Try partial matching (if category contains our search term)
            for category in all_categories:
                if (original_name.lower() in category['name'].lower() or
                    standardized_with_and.lower() in category['name'].lower() or
                    standardized_with_ampersand.lower() in category['name'].lower()):
                    category_ids.append(category['id'])
                    print(f"Found partial match for '{original_name}' → '{category['name']}' with ID: {category['id']}")
                    found = True
                    break
                    
            if found:
                continue
                
            # Try individual words (excluding common words)
            words = original_name.split()
            common_words = ['and', 'or', 'the', 'in', 'on', 'at', 'to', 'for', 'with', 'by', 'of']
            significant_words = [word for word in words if len(word) > 2 and word.lower() not in common_words]
            
            if significant_words:
                print(f"Trying word-by-word search for '{original_name}' with words: {significant_words}")
                
                for word in significant_words:
                    for category in all_categories:
                        if word.lower() in category['name'].lower():
                            category_ids.append(category['id'])
                            print(f"Found word match '{word}' in category '{category['name']}' with ID: {category['id']}")
                            found = True
                            break
                    
                    if found:
                        break
            
            if not found:
                print(f"No category matches found for '{original_name}'")

        # Remove duplicates while preserving order
        unique_ids = []
        for id in category_ids:
            if id not in unique_ids:
                unique_ids.append(id)
                
        return unique_ids

    except Exception as e:
        print(f"Error searching for categories: {e}")
        return []
    
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

def create_wordpress_post_with_details(content_data):
    """
    Create WordPress post with detailed response information.
    Returns a dictionary with success status and additional verification details.
    """
    result = {
        'success': False,
        'post_id': None,
        'post_url': None,
        'error': None,
        'featured_media_verified': False,
        'categories_verified': False
    }
    
    try:
        # Format content with HTML
        formatted_content = ""

        # Replace newlines in the redaction with paragraph tags
        formatted_content += ''.join(f"<p>{para.strip()}</p>" for para in content_data['Redaction'].split("\n") if para.strip())

        # Prepare post data with featured image
        post_data = {
            'title': content_data['Headline'],
            'content': formatted_content,
            'status': 'draft',  # Set to 'draft' for review
        }

        # Add author if provided (already looked up in the main function)
        if content_data.get('author_id'):
            post_data['author'] = content_data['author_id']
            print(f"Setting author ID: {content_data['author_id']}")

        # Add categories if provided (already looked up in the main function)
        if content_data.get('category_ids'):
            post_data['categories'] = content_data['category_ids']
            print(f"Setting categories: {content_data['category_ids']}")

        # Explicitly set featured image if available
        if content_data.get('featured_media_id'):
            post_data['featured_media'] = content_data['featured_media_id']
            print(f"Setting featured image ID: {content_data['featured_media_id']}")

        headers = {
            'Content-Type': 'application/json',
        }

        # Send request to WordPress API
        response = requests.post(
            f'{WP_URL}/wp/v2/posts',
            json=post_data,
            auth=(WP_USER, WP_PASSWORD),
            headers=headers,
            timeout=30
        )

        if response.status_code == 201:
            post_data = response.json()
            result['success'] = True
            result['post_id'] = post_data['id']
            result['post_url'] = post_data.get('link')
            print(f"Successfully created post: {post_data.get('link')}")

            # Verify post details
            verify_response = requests.get(
                f'{WP_URL}/wp/v2/posts/{post_data["id"]}',
                auth=(WP_USER, WP_PASSWORD)
            )
            if verify_response.status_code == 200:
                verify_data = verify_response.json()

                # Verify featured image
                if content_data.get('featured_media_id'):
                    if verify_data.get('featured_media') == content_data['featured_media_id']:
                        result['featured_media_verified'] = True
                        print("Featured image successfully set and verified")
                    else:
                        print("Warning: Featured image may not have been set correctly")

                # Verify categories
                if content_data.get('category_ids'):
                    if set(verify_data.get('categories', [])) == set(content_data['category_ids']):
                        result['categories_verified'] = True
                        print("Categories successfully set and verified")
                    else:
                        print("Warning: Categories may not have been set correctly")
                        
            return result
        else:
            result['error'] = f"HTTP {response.status_code}: {response.text}"
            return result

    except Exception as e:
        result['error'] = str(e)
        return result

def update_online_status(sheet_id, cell_reference):
    """Update the 'Online' checkbox in the Google Sheet to checked."""
    try:
        # For checkboxes, we need to use batchUpdate with cell format
        batch_update_body = {
            'requests': [{
                'updateCells': {
                    'range': {
                        'sheetId': 0,  # Assuming it's the first sheet
                        'startRowIndex': int(cell_reference[1:]) - 1,  # Convert cell ref (e.g., 'C8') to row index
                        'endRowIndex': int(cell_reference[1:]),
                        'startColumnIndex': 3,  # Column D is index 3
                        'endColumnIndex': 4
                    },
                    'rows': [{
                        'values': [{
                            'userEnteredValue': {
                                'boolValue': True
                            }
                        }]
                    }],
                    'fields': 'userEnteredValue'
                }
            }]
        }

        # Execute the update
        result = sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=sheet_id,
            body=batch_update_body
        ).execute()

        print(f"Updated checkbox in cell {cell_reference}")
        return True

    except Exception as e:
        print(f"Failed to update spreadsheet checkbox: {e}")
        return False
    
def get_sheet_id(sheet_url):
    match = re.search(r"/d/([a-zA-Z0-9_-]+)", sheet_url)
    if match:
        return match.group(1)
    else:
        raise ValueError("Invalid Google Sheets URL")

def main(sheet_id):
    """Main function to process eligible posts and summarize by section."""
    successful_posts = []
    failed_posts = []
    request_delay = 3  # Seconds between API requests
    
    # ANSI color codes for terminal output
    GREEN = "\033[92m"
    YELLOW = "\033[93m"
    RED = "\033[91m"
    BLUE = "\033[94m"
    ENDC = "\033[0m"
    BOLD = "\033[1m"
    
    try:
        print(f"{BLUE}{BOLD}Starting processing...{ENDC}")
        # Get eligible rows
        try:
            eligible_rows = get_eligible_rows(sheet_id)
            print(f"{BLUE}Found {len(eligible_rows)} eligible posts\n{ENDC}")
        except Exception as e:
            print(f"{RED}Error getting eligible rows: {e}{ENDC}")
            return

        for row in eligible_rows:
            print(f"\n{BOLD}Processing row {row['row']} (Section: {row['section']}){ENDC}")
            
            # Enhanced post info with detailed status tracking
            post_info = {
                'row': row['row'],
                'headline': 'Unknown',  # Default value
                'authors': row['author_names'],
                'photographer': row.get('photographer_name'),
                'categories': row['categories'],
                'section': row['section'],
                
                # Status tracking fields
                'status': 'Failed',  # Will be updated to 'Success' if post is created
                'post_url': None,
                'post_id': None,
                'image_status': {
                    'has_image': False,
                    'status': 'No image found',
                    'media_id': None
                },
                'category_status': {
                    'requested': len(row['categories']),
                    'applied': 0,
                    'status': 'Not processed'
                },
                'author_status': {
                    'requested': len(row['author_names']),
                    'applied': 0,
                    'primary_author_id': None,
                    'status': 'Not processed'
                },
                'sheet_update_status': 'Not updated',
                'error_details': None
            }

            try:
                # Extract Google Doc ID
                doc_match = re.search(r'/document/d/([a-zA-Z0-9_-]+)', row['doc_url'])
                if not doc_match:
                    raise ValueError('Invalid Google Doc URL')
                doc_id = doc_match.group(1)

                # Parse Google Doc
                sections = parse_google_doc(doc_id)

                # Clean up headline
                original_headline = sections['Headline']
                cleaned_headline = re.sub(r'\s*\bSH:?\b\s*:?\s*', ': ', original_headline, flags=re.IGNORECASE)
                cleaned_headline = ' '.join(cleaned_headline.split()).strip()
                sections['Headline'] = cleaned_headline

                print("Parsed document sections:")
                for key, value in sections.items():
                    print(f"{key}: {value[:50]}...")

                # Update post info with actual headline
                post_info['headline'] = sections['Headline']

                # Handle featured image
                featured_media_id = None
                if sections['Featured image']:
                    image_match = re.search(r'https?://drive\.google\.com/\S+', sections['Featured image'])
                    if image_match:
                        image_url = image_match.group()
                        file_id = extract_file_id(image_url)
                        if file_id:
                            print(f"Downloading image with file ID: {file_id}")
                            image_data = download_image(file_id)
                            if image_data:
                                filename = f"featured_image_{doc_id}.jpg"
                                print(f"Uploading image: {filename}")
                                # Use improved upload function with retries
                                featured_media_id = upload_image_to_wordpress(
                                    image_data,
                                    sections['Cutlines'],
                                    filename,
                                    max_retries=3,
                                    retry_delay=3
                                )
                                if featured_media_id:
                                    post_info['image_status']['has_image'] = True
                                    post_info['image_status']['status'] = 'Uploaded successfully'
                                    post_info['image_status']['media_id'] = featured_media_id
                                    print(f"{GREEN}Successfully uploaded image with ID: {featured_media_id}{ENDC}")
                                else:
                                    post_info['image_status']['status'] = 'Upload failed after retries'
                                    print(f"{YELLOW}Failed to upload image after multiple attempts{ENDC}")
                            else:
                                post_info['image_status']['status'] = 'Download failed'
                                print(f"{YELLOW}Failed to download image data{ENDC}")
                        else:
                            post_info['image_status']['status'] = 'Invalid file ID'
                            print(f"{YELLOW}Failed to extract file ID from image URL{ENDC}")
                    else:
                        post_info['image_status']['status'] = 'No valid URL'
                        print(f"{YELLOW}No valid image URL found in Featured image section{ENDC}")

                # Process author information - ONLY DO THIS ONCE
                author_id = None
                if row['author_names']:
                    author_name = row['author_names'][0]
                    author_id = get_or_create_author_id(author_name)
                    if author_id:
                        post_info['author_status']['primary_author_id'] = author_id
                        post_info['author_status']['applied'] = 1
                        post_info['author_status']['status'] = 'Primary author set'
                        if len(row['author_names']) > 1:
                            post_info['author_status']['status'] += f", {len(row['author_names']) - 1} co-authors need manual addition"
                    else:
                        post_info['author_status']['status'] = 'Author creation failed'

                # Process category information - ONLY DO THIS ONCE
                category_ids = []
                if row['categories']:
                    category_ids = get_category_ids(row['categories'])
                    post_info['category_status']['applied'] = len(category_ids)
                    if category_ids:
                        if len(category_ids) == len(row['categories']):
                            post_info['category_status']['status'] = 'All categories applied'
                        else:
                            post_info['category_status']['status'] = f"{len(category_ids)}/{len(row['categories'])} categories found"
                    else:
                        post_info['category_status']['status'] = 'No categories found'

                # Create post content with pre-looked-up IDs
                post_content = {
                    'Headline': sections['Headline'],
                    'Redaction': sections['Redaction'],
                    'featured_media_id': featured_media_id,
                    'author_id': author_id,           # Pass the already-retrieved author ID
                    'category_ids': category_ids      # Pass the already-retrieved category IDs
                }

                # Create WordPress post with detailed response capture
                post_response = create_wordpress_post_with_details(post_content)
                if post_response['success']:
                    post_info['status'] = 'Success'
                    post_info['post_id'] = post_response['post_id']
                    post_info['post_url'] = post_response['post_url']
                    
                    # Update verification statuses
                    if 'featured_media_verified' in post_response:
                        if post_response['featured_media_verified']:
                            post_info['image_status']['status'] += ' and verified'
                        else:
                            post_info['image_status']['status'] += ' but verification failed'
                            
                    if 'categories_verified' in post_response:
                        if post_response['categories_verified']:
                            post_info['category_status']['status'] += ' and verified'
                        else:
                            post_info['category_status']['status'] += ' but verification failed'
                            
                    # Update spreadsheet status
                    sheet_updated = update_online_status(sheet_id, row['online_cell'])
                    post_info['sheet_update_status'] = 'Updated successfully' if sheet_updated else 'Update failed'
                    
                    successful_posts.append(post_info)
                    # Don't duplicate the success message - it's already printed in create_wordpress_post_with_details
                else:
                    post_info['error_details'] = post_response['error']
                    raise Exception(f"Failed to create WordPress post: {post_response['error']}")
                
                # Add delay between processing posts to avoid rate limiting
                print(f"Waiting {request_delay} seconds before processing next post...")
                time.sleep(request_delay)

            except Exception as e:
                error_message = str(e)
                post_info['error_details'] = error_message
                failed_posts.append(post_info)
                print(f"{RED}Error processing row {row['row']}: {error_message}{ENDC}")
                # Still add delay even after failure
                print(f"Waiting {request_delay} seconds before processing next post...")
                time.sleep(request_delay)
                continue

        # Print detailed summary grouped by section
        print(f"\n{BOLD}{BLUE}" + "="*70)
        print("POSTING SUMMARY BY SECTION")
        print("="*70 + f"{ENDC}")

        # Get all unique sections
        all_sections = set([post['section'] for post in successful_posts + failed_posts])
        
        # Create summary for each section
        for section in sorted(all_sections):
            print(f"\n{BOLD}{BLUE}📌 SECTION: {section}{ENDC}")
            print("-"*70)
            
            # Filter successful posts for this section
            section_successful = [post for post in successful_posts if post['section'] == section]
            if section_successful:
                print(f"\n{GREEN}{BOLD}✅ POSTS CREATED SUCCESSFULLY{ENDC}")
                print("-"*50)
                for post in section_successful:
                    print(f"\n{BOLD}Row {post['row']}: {post['headline']}{ENDC}")
                    print(f"🔗 Post URL: {post['post_url']}")
                    
                    # Author information
                    if len(post['authors']) > 1:
                        print(f"✍️  {BOLD}Authors:{ENDC}")
                        print(f"   Primary author: {post['authors'][0]} (ID: {post['author_status']['primary_author_id']})")
                        print(f"   Co-authors to add manually: {', '.join(post['authors'][1:])}")
                    else:
                        print(f"✍️  {BOLD}Author:{ENDC} {post['authors'][0] if post['authors'] else 'No author specified'}")
                    print(f"    Status: {post['author_status']['status']}")
                    
                    # Category information
                    print(f"🏷️  {BOLD}Categories:{ENDC}")
                    print(f"    Requested ({post['category_status']['requested']}): {', '.join(post['categories'])}")
                    print(f"    Status: {post['category_status']['status']}")
                    
                    # Image information
                    print(f"🖼️  {BOLD}Featured Image:{ENDC}")
                    print(f"    Status: {post['image_status']['status']}")
                    if post['photographer']:
                        print(f"    Photographer: {post['photographer']}")
                        
                    # Spreadsheet update status
                    print(f"📊 {BOLD}Spreadsheet:{ENDC} {post['sheet_update_status']}")
            
            # Filter failed posts for this section
            section_failed = [post for post in failed_posts if post['section'] == section]
            if section_failed:
                print(f"\n{RED}{BOLD}❌ POSTS WITH ERRORS{ENDC}")
                print("-"*50)
                for post in section_failed:
                    print(f"\n{BOLD}Row {post['row']}: {post['headline']}{ENDC}")
                    
                    # Error details
                    print(f"{RED}Error: {post['error_details']}{ENDC}")
                    
                    # Display any progress that was made before failure
                    if post['image_status']['has_image']:
                        print(f"🖼️  Image: {post['image_status']['status']}")
                    
                    if post['author_status']['primary_author_id']:
                        print(f"✍️  Author: {post['author_status']['status']}")
                    
                    if post['category_status']['applied'] > 0:
                        print(f"🏷️  Categories: {post['category_status']['status']}")
                    
                    print(f"{YELLOW}Action needed: Manual posting required{ENDC}")
            
            # Section summary
            print(f"\n→ {BOLD}Section '{section}' summary:{ENDC} {len(section_successful)} posted, {len(section_failed)} failed")

            # Print tabular report for quick reference
            print(f"\n{BOLD}{BLUE}" + "="*70)
            print("QUICK REFERENCE SUMMARY")
            print("="*70 + f"{ENDC}")

            # Print header
            print(f"{BOLD}{'Row':^5} | {'Section':^15} | {'Status':^10} | {'Image':^12} | {'Categories':^30} | {'Authors':^30}{ENDC}")
            print("-"*110)  # Increased width for the longer content

            # Sort by section and row for a cleaner presentation
            all_posts = sorted(successful_posts + failed_posts, key=lambda x: (x['section'], x['row']))

            for post in all_posts:
                # Format status color based on success/failure
                status_color = GREEN if post['status'] == 'Success' else RED
                
                # Format image status color
                if 'success' in post['image_status']['status'].lower():
                    img_color = GREEN
                elif 'no image' in post['image_status']['status'].lower():
                    img_color = YELLOW
                else:
                    img_color = RED
                    
                # Format actual category list instead of status
                categories_display = ', '.join(post['categories'][:3])  # First 3 categories
                if len(post['categories']) > 3:
                    categories_display += f", +{len(post['categories'])-3} more"
                
                # Format actual author list instead of status
                authors_display = ', '.join(post['authors'][:3])  # First 3 authors
                if len(post['authors']) > 3:
                    authors_display += f", +{len(post['authors'])-3} more"
                
                # Print row with colors
                print(f"{post['row']:^5} | {post['section'][:15]:^15} | {status_color}{post['status']:^10}{ENDC} | "
                    f"{img_color}{post['image_status']['status'][:12]:^12}{ENDC} | "
                    f"{categories_display[:30]:30} | "
                    f"{authors_display[:30]:30}")

        # Overall summary with percentages
        total_posts = len(successful_posts) + len(failed_posts)
        success_rate = (len(successful_posts) / total_posts * 100) if total_posts > 0 else 0
        
        # Image statistics
        images_with_success = sum(1 for post in successful_posts if post['image_status']['has_image'])
        image_success_rate = (images_with_success / len(successful_posts) * 100) if successful_posts else 0
        
        # Category statistics
        categories_requested = sum(post['category_status']['requested'] for post in successful_posts)
        categories_applied = sum(post['category_status']['applied'] for post in successful_posts)
        category_success_rate = (categories_applied / categories_requested * 100) if categories_requested > 0 else 0
        
        print(f"\n{BOLD}{BLUE}" + "="*70)
        print("OVERALL SUMMARY")
        print("="*70 + f"{ENDC}")
        print(f"{BOLD}Total sections:{ENDC} {len(all_sections)}")
        print(f"{BOLD}Total posts processed:{ENDC} {total_posts}")
        print(f"{BOLD}Successfully posted:{ENDC} {len(successful_posts)} ({success_rate:.1f}%)")
        print(f"{BOLD}Failed posts:{ENDC} {len(failed_posts)} ({100-success_rate:.1f}%)")
        print(f"{BOLD}Image success rate:{ENDC} {images_with_success}/{len(successful_posts)} posts with images ({image_success_rate:.1f}%)")
        print(f"{BOLD}Category success rate:{ENDC} {categories_applied}/{categories_requested} categories applied ({category_success_rate:.1f}%)")
        print(f"{BLUE}{BOLD}" + "="*70 + f"{ENDC}\n")

    except Exception as e:
        print(f"{RED}{BOLD}Fatal error: {e}{ENDC}")

if __name__ == '__main__':
    # The Sheet ID is the long string in the URL of the Google Sheet
    # In "https://docs.google.com/spreadsheets/d/1RFuJl1VAFeeCmJdgtsFvH0ZTV3irataUe2oyvWzwnA0/edit?gid=0#gid=0", the sheet ID is between d/ and /edit
    # Ask user for the spreadsheet URL
    sheet_url = input("Enter Google Sheets URL: ").strip()

    try:
        sheetid = get_sheet_id(sheet_url)
        print("Extracted Sheet ID:", sheetid)
    except ValueError as e:
        print(e)

    main(sheetid)