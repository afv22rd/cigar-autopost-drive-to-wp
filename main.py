import re
import requests
import os
import io
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
    - 'Online' (Column C) is NOT checked
    Handles formatted text with hyperlinks in Photographer column.
    """
    sheet = sheets_service.spreadsheets().get(
        spreadsheetId=sheet_id,
        includeGridData=True,
        fields="sheets(data(rowData(values(formattedValue,textFormatRuns,userEnteredFormat.backgroundColor))))"
    ).execute()
    rows = sheet['sheets'][0]['data'][0]['rowData']

    eligible_rows = []
    for row_idx, row in enumerate(rows[7:]):  # Skip header row
        try:
            # Get cell values
            ready_to_post = row['values'][1].get('formattedValue', '').upper() == 'TRUE'  # Column B
            is_online = row['values'][2].get('formattedValue', '').upper() == 'TRUE'      # Column C

            # Check eligibility criteria
            if ready_to_post and not is_online:
                story_cell = row['values'][3]  # Column D (Story)

                # Extract hyperlink from Story cell
                hyperlink = None

                # First, try to get hyperlink from textFormatRuns
                for run in story_cell.get('textFormatRuns', []):
                    if 'link' in run.get('format', {}):
                        hyperlink = run['format']['link']['uri']
                        break

                # Fallback: if no hyperlink in textFormatRuns, check if the formattedValue itself is a URL
                if not hyperlink:
                    url_match = re.search(r'https?://[^\s]+', story_cell.get('formattedValue', ''))
                    if url_match:
                        hyperlink = url_match.group()

                if hyperlink:
                    # Check for photographer link (Column M)
                    photographer_link = None
                    photographer_name = None
                    if len(row['values']) > 12:  # Ensure column M exists
                        photographer_cell = row['values'][12]
                        # Get the photographer's name (formatted text)
                        photographer_name = photographer_cell.get('formattedValue', '')

                    # Log photographer information
                    if photographer_name:
                        print(f"Row {row_idx + 8}: Photography found - '{photographer_name}'")
                    else:
                        print(f"Row {row_idx + 8}: No photography information found")

                    # Add row to eligible list regardless of photographer link
                    eligible_rows.append({
                        'row': row_idx + 8,  # Adjust for header
                        'doc_url': hyperlink,
                        'photographer_link': photographer_link,
                        'photographer_name': photographer_name,
                        'online_cell': f"C{row_idx + 8}"  # Column C (Online)
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
        'Author(s)': '',
        'Categories': ''
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

def get_author_id(author_names):
    """
    Search WordPress users by name and return their user ID.
    Now handles multiple authors, returns first author's ID and logs co-authors.
    """
    try:
        # Split authors by comma and clean whitespace
        authors = [name.strip() for name in author_names.split(',')]

        # Log if multiple authors detected
        if len(authors) > 1:
            print(f"WARNING: Multiple authors detected: {author_names}")
            print(f"Using primary author '{authors[0]}'. Please manually add these co-authors: {', '.join(authors[1:])}")

        # Get first author's ID
        headers = {
            'Authorization': f'Basic {b64encode(f"{WP_USER}:{WP_PASSWORD}".encode()).decode()}'
        }

        # Use the specific users endpoint
        users_endpoint = f'{WP_URL}/wp/v2/users'

        # Search for the primary author
        params = {'search': authors[0]}
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
                    if user['name'].lower() == authors[0].lower():
                        print(f"Found exact match for primary author '{authors[0]}' with ID: {user['id']}")
                        return user['id']

                # If no exact match, return first result
                print(f"Found partial match for primary author '{authors[0]}' with ID: {users[0]['id']}")
                return users[0]['id']

        print(f"Primary author '{authors[0]}' not found. Response: {response.text}")
        return None

    except Exception as e:
        print(f"Error searching for author: {e}")
        return None

def get_category_ids(categories_string):
    """Search WordPress categories by name and return their IDs."""
    try:
        headers = {
            'Authorization': f'Basic {b64encode(f"{WP_USER}:{WP_PASSWORD}".encode()).decode()}'
        }

        # Split and clean category names
        names_list = [name.strip() for name in categories_string.split(',')]
        category_ids = []

        # Use the categories endpoint
        categories_endpoint = f'{WP_URL}/wp/v2/categories'

        # Search for each category individually
        for name in names_list:
            params = {'search': name}
            response = requests.get(
                categories_endpoint,
                headers=headers,
                params=params,
                timeout=10
            )

            if response.status_code == 200:
                categories = response.json()
                if categories:
                    # Look for exact match (case-insensitive)
                    found = False
                    for category in categories:
                        if category['name'].lower() == name.lower():
                            category_ids.append(category['id'])
                            print(f"Found category '{name}' with ID: {category['id']}")
                            found = True
                            break

                    if not found:
                        print(f"Category '{name}' not found")
                else:
                    print(f"No categories found matching '{name}'")
            else:
                print(f"Failed to search for category '{name}'. Response: {response.text}")

        return category_ids

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

def upload_image_to_wordpress(image_data, caption, filename):
    """Upload image to WordPress media library with improved error handling."""
    if not image_data:
        print("No image data provided")
        return None

    mime_type = mimetypes.guess_type(filename)[0] or 'image/jpeg'

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
            return None

    except Exception as e:
        print(f"Image upload failed: {str(e)}")
        return None

def create_wordpress_post(content_data):
    """Create WordPress post with formatted content and featured image."""
    try:
        # Format content with HTML
        formatted_content = ""

        # Replace newlines in the redaction with paragraph tags
        formatted_content += ''.join(f"<p>{para.strip()}</p>" for para in content_data['Redaction'].split("\n") if para.strip())

        # Get Author ID
        author_id = None;
        if content_data['Author(s)']:
            # Clean up author name (remove any extra whitespace or newlines)
            author_name = content_data['Author(s)'].strip()
            author_id = get_author_id(author_name)

        # Get category IDs
        category_ids = []
        if content_data.get('Categories'):
            category_ids = get_category_ids(content_data['Categories'])
            if not category_ids:
                print("Warning: No valid categories found")

        # Prepare post data with featured image
        post_data = {
            'title': content_data['Headline'],
            'content': formatted_content,
            'status': 'draft',  # Set to 'draft' for review
        }

        # Add author if found
        if author_id:
            post_data['author'] = author_id
            print(f"Setting author ID: {author_id}")
        else:
            print(f"Warning: Could not find author ID for '{content_data['Author(s)']}'")

        # Add categories if found
        if category_ids:
            post_data['categories'] = category_ids
            print(f"Setting categories: {category_ids}")
        else:
            print(f"Warning: Could not find categories for '{content_data['Categories']}'")

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
                        print("Featured image successfully set")
                    else:
                        print("Warning: Featured image may not have been set correctly")

                # Verify categories
                if category_ids:
                    if set(verify_data.get('categories', [])) == set(category_ids):
                        print("Categories successfully set")
                    else:
                        print("Warning: Categories may not have been set correctly")

            return True
        else:
            print(f"Post creation failed: {response.text}")
            return False

    except Exception as e:
        print(f"Post creation error: {e}")
        return False

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
                        'startColumnIndex': 2,  # Column C is index 2 (0-based)
                        'endColumnIndex': 3
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

def main(sheet_id):
    """Main function to process eligible posts."""
    successful_posts = []
    failed_posts = []
    try:
        print("Starting processing...")
        # Get eligible rows
        try:
          eligible_rows = get_eligible_rows(sheet_id)
          print(f"Found {len(eligible_rows)} eligible posts\n")
        except Exception as e:
          print(f"Error getting eligible rows: {e}")
          return

        for row in eligible_rows:
            try:
                print(f"\nProcessing row {row['row']}")

                # Extract Google Doc ID
                doc_match = re.search(r'/document/d/([a-zA-Z0-9_-]+)', row['doc_url'])
                if not doc_match:
                    failed_posts.append({
                        'row': row['row'],
                        'error': 'Invalid Google Doc URL',
                        'headline': 'N/A'
                    })
                    print("Invalid Google Doc URL")
                    continue
                doc_id = doc_match.group(1)

                # Parse Google Doc
                sections = parse_google_doc(doc_id)
                print("Parsed document sections:")
                for key, value in sections.items():
                    print(f"{key}: {value[:50]}...")

                # Track post information
                post_info = {
                    'row': row['row'],
                    'headline': sections['Headline'],
                    'has_image': False,
                    'authors': [author.strip() for author in sections['Author(s)'].split(',')],
                    'photographer': row.get('photographer_name')
                }

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
                                featured_media_id = upload_image_to_wordpress(
                                    image_data,
                                    sections['Cutlines'],
                                    filename
                                )
                                if featured_media_id:
                                    post_info['has_image'] = True
                                    print(f"Successfully uploaded image with ID: {featured_media_id}")
                            else:
                                print("Failed to download image data")
                        else:
                            print("Failed to extract file ID from image URL")
                    else:
                        print("No valid image URL found in Featured image section")

                # Create post content
                post_content = {
                    'Headline': sections['Headline'],
                    'Redaction': sections['Redaction'],
                    'Author(s)': sections['Author(s)'],
                    'Categories': sections['Categories'],
                    'featured_media_id': featured_media_id
                }

                # Create WordPress post
                if create_wordpress_post(post_content):
                    update_online_status(sheet_id, row['online_cell'])
                    successful_posts.append(post_info)
                else:
                    failed_posts.append({
                        'row': row['row'],
                        'headline': sections['Headline'],
                        'error': 'Failed to create WordPress post'
                    })

            except Exception as e:
                failed_posts.append({
                    'row': row['row'],
                    'headline': sections.get('Headline', 'Unknown'),
                    'error': str(e)
                })
                print(f"Error processing row {row['row']}: {e}")
                continue

        # Print summary
        print("\n" + "="*50)
        print("POSTING SUMMARY")
        print("="*50)

        # Successful posts
        print("\n✅ POSTS CREATED SUCCESSFULLY")
        print("-"*30)
        for post in successful_posts:
            print(f"\nRow {post['row']}: {post['headline']}")
            print(f"📸 Photography: {'Yes - ' + post['photographer'] if post['photographer'] else 'No'}")
            if len(post['authors']) > 1:
                print(f"✍️  Multiple authors detected:")
                print(f"   Primary author: {post['authors'][0]}")
                print(f"   Co-authors to add: {', '.join(post['authors'][1:])}")
            else:
                print(f"✍️  Author: {post['authors'][0]}")

        # Failed posts
        if failed_posts:
            print("\n❌ POSTS WITH ERRORS")
            print("-"*30)
            for post in failed_posts:
                print(f"\nRow {post['row']}: {post['headline']}")
                print(f"Error: {post['error']}")
                print(f"Action needed: Manual posting required")

        print("\n" + "="*50)
        print(f"Total posts processed: {len(successful_posts) + len(failed_posts)}")
        print(f"Successfully posted: {len(successful_posts)}")
        print(f"Failed posts: {len(failed_posts)}")
        print("="*50 + "\n")

    except Exception as e:
        print(f"Fatal error: {e}")

if __name__ == '__main__':
    main('1RFuJl1VAFeeCmJdgtsFvH0ZTV3irataUe2oyvWzwnA0')