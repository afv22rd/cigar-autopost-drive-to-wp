import re
import requests
import os
import io
import time
import keyboard
import termios, tty, sys
import platform
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

            # Get Image URL (Column N)
            image_url = None
            if len(values) > 13:  # Column N is index 13
                image_cell = values[13]
                print(f"Row {actual_row_num}: Analyzing image cell (Column N):")
                print(f"  Cell content: {image_cell}")

                # Method 1: Try to get URL from textFormatRuns
                if 'textFormatRuns' in image_cell:
                    print(f"  textFormatRuns found in image cell")
                    for run in image_cell['textFormatRuns']:
                        if 'format' in run and 'link' in run['format']:
                            image_url = run['format']['link']['uri']
                            print(f"  Found image URL from textFormatRuns: {image_url}")
                            break

                # Method 2: Try to get URL from hyperlink property
                if not image_url and 'hyperlink' in image_cell:
                    image_url = image_cell['hyperlink']
                    print(f"  Found image URL from hyperlink property: {image_url}")

                # Method 3: Look for URL patterns in text
                if not image_url and 'formattedValue' in image_cell:
                    url_match = re.search(r'https?://[^\s]+', image_cell['formattedValue'])
                    if url_match:
                        image_url = url_match.group()
                        print(f"  Found image URL from text pattern: {image_url}")
            else:
                print(f"Row {actual_row_num}: Story has no featured image.")

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
                
            # Get Photographer info (Column P)
            photographer_name = None
            if len(values) > 15:  # Column P is index 15
                photographer_cell = values[15]
                if photographer_cell and 'formattedValue' in photographer_cell:
                    photographer_name = photographer_cell.get('formattedValue', '').strip()
                    if photographer_name:
                        print(f"Row {actual_row_num}: Found photographer: {photographer_name}")

            # Get Headlines document URL (Column P)
            headlines_url = None
            if len(values) > 15:  # Column P is index 15
                headlines_cell = values[15]
                print(f"Row {actual_row_num}: Analyzing headlines cell (Column P):")
                
                # Method 1: Try to get URL from textFormatRuns
                if 'textFormatRuns' in headlines_cell:
                    for run in headlines_cell['textFormatRuns']:
                        if 'format' in run and 'link' in run['format']:
                            headlines_url = run['format']['link']['uri']
                            print(f"  Found headlines URL from textFormatRuns: {headlines_url}")
                            break

                # Method 2: Try to get URL from hyperlink property
                if not headlines_url and 'hyperlink' in headlines_cell:
                    headlines_url = headlines_cell['hyperlink']
                    print(f"  Found headlines URL from hyperlink property: {headlines_url}")

                # Method 3: Look for URL patterns in text
                if not headlines_url and 'formattedValue' in headlines_cell:
                    url_match = re.search(r'https?://[^\s]+', headlines_cell['formattedValue'])
                    if url_match:
                        headlines_url = url_match.group()
                        print(f"  Found headlines URL from text pattern: {headlines_url}")
            
            # Get Cutlines document URL (Column Q)
            cutlines_url = None
            if len(values) > 16:  # Column Q is index 16
                cutlines_cell = values[16]
                print(f"Row {actual_row_num}: Analyzing cutlines cell (Column Q):")
                
                # Method 1: Try to get URL from textFormatRuns
                if 'textFormatRuns' in cutlines_cell:
                    for run in cutlines_cell['textFormatRuns']:
                        if 'format' in run and 'link' in run['format']:
                            cutlines_url = run['format']['link']['uri']
                            print(f"  Found cutlines URL from textFormatRuns: {cutlines_url}")
                            break

                # Method 2: Try to get URL from hyperlink property
                if not cutlines_url and 'hyperlink' in cutlines_cell:
                    cutlines_url = cutlines_cell['hyperlink']
                    print(f"  Found cutlines URL from hyperlink property: {cutlines_url}")

                # Method 3: Look for URL patterns in text
                if not cutlines_url and 'formattedValue' in cutlines_cell:
                    url_match = re.search(r'https?://[^\s]+', cutlines_cell['formattedValue'])
                    if url_match:
                        cutlines_url = url_match.group()
                        print(f"  Found cutlines URL from text pattern: {cutlines_url}")

            # Add to eligible rows
            print(f"Row {actual_row_num}: Adding to eligible rows (Section: {current_section})")
            eligible_rows.append({
                'row': actual_row_num,
                'doc_url': story_url,
                'image_url': image_url,
                'headlines_url': headlines_url,
                'cutlines_url': cutlines_url,
                'author_names': author_names,
                'categories': categories,
                'photographer_name': photographer_name,
                'online_cell': f"D{actual_row_num}",
                'section': current_section
            })

        except Exception as e:
            print(f"Error processing row {row_idx + 8}: {str(e)}")
            continue

    return eligible_rows

def parse_redaction_doc(doc_id):
    """
    Parse redaction document with interactive line selection.
    Shows first 9 lines and lets user select where redaction starts.
    """
    # ANSI color codes for terminal output
    GREEN = "\033[92m"
    YELLOW = "\033[93m"
    BLUE = "\033[94m"
    ENDC = "\033[0m"
    BOLD = "\033[1m"
    
    try:
        # Fetch the Google Doc content
        doc = docs_service.documents().get(documentId=doc_id).execute()
        content = doc['body']['content']
        
        # Extract all lines from the document
        all_lines = []
        
        for element in content:
            if 'paragraph' in element:
                elements = element['paragraph']['elements']
                text = ''.join([e.get('textRun', {}).get('content', '') for e in elements])
                all_lines.append(text.strip())
        
        # Remove any empty lines at the beginning
        while all_lines and not all_lines[0]:
            all_lines.pop(0)
            
        if not all_lines:
            print(f"{YELLOW}Warning: Document appears to be empty{ENDC}")
            return ""
        
        # Display first 9 lines (or fewer if document has fewer lines)
        display_lines = min(9, len(all_lines))
        
        print(f"\n{BLUE}{BOLD}First {display_lines} lines of the redaction document:{ENDC}")
        for i in range(display_lines):
            # Truncate long lines
            display_text = all_lines[i][:100] + "..." if len(all_lines[i]) > 100 else all_lines[i]
            print(f"{i+1}. {display_text}")
        
        # Prompt for redaction start line
        print(f"\n{BOLD}Where does the redaction start?{ENDC} (default: line 4)")
        print(f"Press {GREEN}SPACEBAR{ENDC} for default (line 4) or enter a line number:")
        
        # Get user input
        user_input = ""
        start_line = 4  # Default value
        
        while True:
            char = get_single_key()
            if char == ' ':  # Spacebar = default
                print(f"Using default starting line: {start_line}")
                break
            elif char in ['\r', '\n']:  # Enter = use current input
                if user_input.isdigit() and 1 <= int(user_input) <= len(all_lines):
                    start_line = int(user_input)
                    print(f"Using starting line: {start_line}")
                    break
                elif user_input == "":
                    print(f"Using default starting line: {start_line}")
                    break
                else:
                    print(f"{YELLOW}Invalid input. Please enter a valid line number (1-{len(all_lines)}):{ENDC}")
                    user_input = ""
            elif char.isdigit():
                user_input += char
                print(char, end="", flush=True)
            elif char in ['\b', '\x08', '\x7f']:  # Backspace
                if user_input:
                    user_input = user_input[:-1]
                    print("\b \b", end="", flush=True)
            
        # Create redaction from selected line to end
        start_idx = start_line - 1  # Convert to 0-based index
        redaction = '\n'.join([line for line in all_lines[start_idx:] if line])
        
        # For preview purposes, show the first few lines of redaction
        preview_lines = redaction.split('\n')[:3]
        preview_text = '\n'.join(preview_lines)
        if len(preview_lines) < len(redaction.split('\n')):
            preview_text += "\n..."
        
        print(f"\n{BOLD}Redaction content:{ENDC}")
        print(f"{preview_text}")
        
        return redaction
        
    except Exception as e:
        print(f"Error parsing redaction document: {str(e)}")
        return f"Error parsing document: {str(e)}"
    
def parse_headlines_doc(doc_id):
    """
    Parse headlines document and return a list of headline options.
    Format expected: "**SLUG - **HEADLINE"
    """
    # ANSI color codes for terminal output
    YELLOW = "\033[93m"
    ENDC = "\033[0m"
    
    try:
        # Fetch the Google Doc content
        doc = docs_service.documents().get(documentId=doc_id).execute()
        content = doc['body']['content']
        
        # Extract all lines from the document
        all_lines = []
        
        for element in content:
            if 'paragraph' in element:
                elements = element['paragraph']['elements']
                text = ''.join([e.get('textRun', {}).get('content', '') for e in elements])
                if text.strip():  # Only append non-empty lines
                    all_lines.append(text.strip())
        
        # Extract headlines using the specific pattern
        headlines = []
        pattern = re.compile(r'\*\*(.*?)\s*-\s*\*\*(.*?)(?=$|\*\*)', re.DOTALL)
        
        # Process each line in the document
        for line in all_lines:
            # Find all matches in this line
            matches = pattern.finditer(line)
            for match in matches:
                slug = match.group(1).strip()
                headline = match.group(2).strip()
                headlines.append({
                    'slug': slug,
                    'headline': headline,
                    'original': f"**{slug} - **{headline}"
                })
        
        # Also check for pattern across the whole document text
        full_text = ' '.join(all_lines)
        matches = pattern.finditer(full_text)
        for match in matches:
            slug = match.group(1).strip()
            headline = match.group(2).strip()
            found = False
            # Check if this headline is already in our list
            for existing in headlines:
                if existing['slug'] == slug and existing['headline'] == headline:
                    found = True
                    break
            if not found:
                headlines.append({
                    'slug': slug,
                    'headline': headline,
                    'original': f"**{slug} - **{headline}"
                })
        
        # Clean up headlines - fix "SH:" pattern
        for headline_item in headlines:
            # Replace "SH:" with ": " pattern
            headline_item['headline'] = re.sub(r'\s*SH:\s*', ': ', headline_item['headline'])
            # Remove extra colons if they exist after replacing SH:
            headline_item['headline'] = re.sub(r':\s*:', ':', headline_item['headline'])
        
        print(f"Found {len(headlines)} potential headlines in document")
        return headlines
        
    except Exception as e:
        print(f"Error parsing headlines document: {str(e)}")
        return []

def parse_cutlines_doc(doc_id):
    """
    Parse cutlines document and return a list of cutline options.
    Format expected: "- SLUG: cutline text"
    """
    # ANSI color codes for terminal output
    YELLOW = "\033[93m"
    ENDC = "\033[0m"
    
    try:
        # Fetch the Google Doc content
        doc = docs_service.documents().get(documentId=doc_id).execute()
        content = doc['body']['content']
        
        # Extract all lines from the document
        all_lines = []
        
        for element in content:
            if 'paragraph' in element:
                elements = element['paragraph']['elements']
                text = ''.join([e.get('textRun', {}).get('content', '') for e in elements])
                if text.strip():  # Only append non-empty lines
                    all_lines.append(text.strip())
        
        # Extract cutlines which typically start with "- " or directly with a slug followed by ":"
        cutlines = []
        pattern = re.compile(r'-?\s*(.*?):\s*(.*?)(?=$|-\s*\w+:|$)', re.DOTALL)
        
        # Process each line in the document
        for line in all_lines:
            # Find all matches in this line
            matches = pattern.finditer(line)
            for match in matches:
                slug = match.group(1).strip()
                cutline = match.group(2).strip()
                if cutline:  # Only add if there's actual cutline text
                    cutlines.append({
                        'slug': slug,
                        'cutline': cutline,
                        'original': f"{slug}: {cutline}"
                    })
        
        # Also scan multi-line cutlines
        full_text = '\n'.join(all_lines)
        matches = pattern.finditer(full_text)
        for match in matches:
            slug = match.group(1).strip()
            cutline = match.group(2).strip()
            if cutline:
                found = False
                # Check if this cutline is already in our list
                for existing in cutlines:
                    if existing['slug'] == slug and existing['cutline'] == cutline:
                        found = True
                        break
                if not found:
                    cutlines.append({
                        'slug': slug,
                        'cutline': cutline,
                        'original': f"{slug}: {cutline}"
                    })
        
        print(f"Found {len(cutlines)} potential cutlines in document")
        return cutlines
        
    except Exception as e:
        print(f"Error parsing cutlines document: {str(e)}")
        return []

def select_headline_interactively(headlines, row_info, redaction_preview):
    """
    Present headline options to user for interactive selection.
    Returns the selected headline text.
    """
    # ANSI color codes for terminal output
    GREEN = "\033[92m"
    YELLOW = "\033[93m"
    BLUE = "\033[94m"
    ENDC = "\033[0m"
    BOLD = "\033[1m"
    
    # If no headlines found
    if not headlines:
        print(f"{YELLOW}No headline options found. Please enter a headline manually:{ENDC}")
        return input("Headline: ").strip()
    
    print(f"\n{BLUE}{BOLD}Processing row {row_info['row']} (Section: {row_info['section']}){ENDC}")
    print(f"\n{BOLD}Redaction preview:{ENDC}")
    print(f"{redaction_preview[:150]}...")
    
    print(f"\n{BOLD}What is the headline of this post?{ENDC}")
    
    # Assign letters A, B, C, etc. to each headline
    choices = {}
    for idx, headline in enumerate(headlines):
        letter = chr(65 + idx)  # A=65, B=66, etc.
        if idx < 26:  # Only support up to 26 options (A-Z)
            choices[letter] = headline
            print(f"{BOLD}{letter}. {headline['original']}{ENDC}")
    
    print(f"\n{YELLOW}Enter letter choice (A-{chr(64 + len(choices))}) or type a custom headline:{ENDC}")
    
    user_input = input("> ").strip()
    
    # Check if the input is a valid letter choice
    if user_input.upper() in choices:
        selected_headline = choices[user_input.upper()]
        print(f"{GREEN}Selected: {selected_headline['headline']}{ENDC}")
        return selected_headline['headline']
    else:
        # Treat input as custom headline
        print(f"{GREEN}Using custom headline: {user_input}{ENDC}")
        return user_input

def select_cutline_interactively(cutlines, headline):
    """
    Present cutline options to user for interactive selection.
    Returns the selected cutline text.
    """
    # ANSI color codes for terminal output
    GREEN = "\033[92m"
    YELLOW = "\033[93m"
    BLUE = "\033[94m"
    ENDC = "\033[0m"
    BOLD = "\033[1m"
    
    # If no cutlines found or no image
    if not cutlines:
        print(f"{YELLOW}No cutline options found. Enter a cutline or press Enter to skip:{ENDC}")
        return input("Cutline: ").strip()
    
    print(f"\n{BOLD}What is the cutline for the featured image?{ENDC}")
    print(f"{BLUE}(For headline: {headline}){ENDC}")
    
    # Assign letters A, B, C, etc. to each cutline
    choices = {}
    for idx, cutline in enumerate(cutlines):
        letter = chr(65 + idx)  # A=65, B=66, etc.
        if idx < 26:  # Only support up to 26 options (A-Z)
            choices[letter] = cutline
            print(f"{BOLD}{letter}. {cutline['original']}{ENDC}")
    
    print(f"\n{YELLOW}Enter letter choice (A-{chr(64 + len(choices))}) or type a custom cutline or press Enter to skip:{ENDC}")
    
    user_input = input("> ").strip()
    
    if not user_input:
        print(f"{YELLOW}Skipping cutline.{ENDC}")
        return ""
    
    # Check if the input is a valid letter choice
    if user_input.upper() in choices:
        selected_cutline = choices[user_input.upper()]
        print(f"{GREEN}Selected: {selected_cutline['cutline']}{ENDC}")
        return selected_cutline['cutline']
    else:
        # Treat input as custom cutline
        print(f"{GREEN}Using custom cutline: {user_input}{ENDC}")
        return user_input

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
    # ANSI color codes for terminal output
    GREEN = "\033[92m"
    YELLOW = "\033[93m"
    RED = "\033[91m"
    BLUE = "\033[94m"
    ENDC = "\033[0m"
    BOLD = "\033[1m"
    
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

def create_wordpress_post_with_details(content_data, status='draft'):
    """
    Create WordPress post with detailed response information.
    Returns a dictionary with success status and additional verification details.
    
    Parameters:
    - content_data: Dictionary containing post content
    - status: Post status ('draft' or 'publish')
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
            'status': status,  # Use provided status
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
            print(f"Successfully created post as '{status}': {post_data.get('link')}")

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
    
def get_single_key():
    """Get a single keypress from the user, cross-platform."""
    # Unix implementation
    fd = sys.stdin.fileno()
    old_settings = termios.tcgetattr(fd)
    try:
        tty.setraw(fd)
        ch = sys.stdin.read(1)
    finally:
        termios.tcsetattr(fd, termios.TCSADRAIN, old_settings)
    return ch

def display_post_details(sections, row, featured_media_available=False, image_source="Column N from spreadsheet"):
    """Display post details for review in a formatted way."""
    # ANSI color codes for terminal output
    GREEN = "\033[92m"
    YELLOW = "\033[93m"
    BLUE = "\033[94m"
    ENDC = "\033[0m"
    BOLD = "\033[1m"
    RED = "\033[91m"
    ORANGE = "\033[33m"
    
    print(f"\n{BOLD}{BLUE}" + "="*70)
    print(f"POST REVIEW - ROW {row['row']} - SECTION: {row['section']}")
    print("="*70 + f"{ENDC}")
    
    # Display headline
    print(f"\n{BOLD}Headline:{ENDC}")
    print(f"{sections['Headline']}")
    
    # Display authors
    print(f"\n{BOLD}Authors:{ENDC}")
    if row['author_names']:
        for i, author in enumerate(row['author_names']):
            print(f"  {'Primary: ' if i == 0 else 'Co-author: '}{author}")
    else:
        print(f"{YELLOW}  No authors specified{ENDC}")
    
    # Display categories
    print(f"\n{BOLD}Categories:{ENDC}")
    if row['categories']:
        for category in row['categories']:
            print(f"  {category}")
    else:
        print(f"{YELLOW}  No categories specified{ENDC}")
    
    # Display featured image status
    print(f"\n{BOLD}Featured Image:{ENDC}")
    if featured_media_available:
        print(f"{GREEN}  Image available{ENDC}")
        
        # Display image source - now using the parameter
        print(f"  Source: {image_source}")
        
        # Display cutlines if available
        if sections.get('Cutlines'):
            print(f"\n{BOLD}Cutlines:{ENDC}")
            print(f"  {sections['Cutlines']}")
        
        # Display photographer if available
        if row.get('photographer_name'):
            print(f"\n{BOLD}Photographer:{ENDC}")
            print(f"  {row['photographer_name']}")
    else:
        print(f"{YELLOW}  No image available{ENDC}")
    
    # Display redaction (content)
    print(f"\n{BOLD}Content:{ENDC}")
    redaction_lines = sections['Redaction'].split('\n')
    
    # Only show first 5 lines and indicate if there's more
    for i, line in enumerate(redaction_lines[:5]):
        if line.strip():
            print(f"  {line[:100]}{'...' if len(line) > 100 else ''}")
    
    if len(redaction_lines) > 5:
        print(f"  {YELLOW}... and {len(redaction_lines) - 5} more lines ...{ENDC}")
    
    # Display commands
    print(f"\n{BOLD}{BLUE}" + "-"*70)
    print("ACTIONS:")
    print(f"{GREEN}[ENTER]{ENDC} Publish post and continue")
    print(f"{YELLOW}[⟸ BACKSPACE]{ENDC} Create as draft and continue")
    print(f"{BLUE}[SPACEBAR]{ENDC} Skip this post and continue")
    print(f"{RED}[ESC]{ENDC} Exit program")
    print("-"*70 + f"{ENDC}")

def main(sheet_id):
    """Main function to process eligible posts with interactive keyboard controls."""
    successful_posts = []
    failed_posts = []
    skipped_posts = []  # New list to track skipped posts

    # Storage for headline and cutline options
    headlines_cache = []
    cutlines_cache = []
    
    # ANSI color codes for terminal output
    GREEN = "\033[92m"
    YELLOW = "\033[93m"
    RED = "\033[91m"
    BLUE = "\033[94m"
    ENDC = "\033[0m"
    BOLD = "\033[1m"
    ORANGE = "\033[33m"
    
    try:
        print(f"{BLUE}{BOLD}Starting interactive processing...{ENDC}")
        # Get eligible rows
        try:
            eligible_rows = get_eligible_rows(sheet_id)
            print(f"{BLUE}Found {len(eligible_rows)} eligible posts\n{ENDC}")
        except Exception as e:
            print(f"{RED}Error getting eligible rows: {e}{ENDC}")
            return
        
        # Parse headlines document (from first row's column P)
        if eligible_rows:
            first_row = eligible_rows[0]
            if first_row.get('headlines_url'):
                print(f"{BLUE}Parsing headlines document...{ENDC}")
                headlines_doc_match = re.search(r'/document/d/([a-zA-Z0-9_-]+)', first_row['headlines_url'])
                if headlines_doc_match:
                    headlines_doc_id = headlines_doc_match.group(1)
                    headlines_cache = parse_headlines_doc(headlines_doc_id)
                else:
                    print(f"{YELLOW}Invalid headlines document URL format.{ENDC}")
            else:
                print(f"{YELLOW}No headlines document URL found.{ENDC}")
                
            # Parse cutlines document (from first row's column Q)
            if first_row.get('cutlines_url'):
                print(f"{BLUE}Parsing cutlines document...{ENDC}")
                cutlines_doc_match = re.search(r'/document/d/([a-zA-Z0-9_-]+)', first_row['cutlines_url'])
                if cutlines_doc_match:
                    cutlines_doc_id = cutlines_doc_match.group(1)
                    cutlines_cache = parse_cutlines_doc(cutlines_doc_id)
                else:
                    print(f"{YELLOW}Invalid cutlines document URL format.{ENDC}")
            else:
                print(f"{YELLOW}No cutlines document URL found.{ENDC}")

        for row in eligible_rows:
            print(f"\n{BOLD}Loading row {row['row']} (Section: {row['section']}){ENDC}")
            
            # Enhanced post info with detailed status tracking
            post_info = {
                'row': row['row'],
                'headline': 'Unknown',  # Default value
                'authors': row['author_names'],
                'photographer': row.get('photographer_name'),
                'categories': row['categories'],
                'section': row['section'],
                
                # Status tracking fields
                'status': 'Skipped',  # Default is now 'Skipped'
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
                # Extract Google Doc ID for the redaction document (Column E)
                doc_match = re.search(r'/document/d/([a-zA-Z0-9_-]+)', row['doc_url'])
                if not doc_match:
                    raise ValueError('Invalid Google Doc URL for redaction')
                doc_id = doc_match.group(1)

                # Parse redaction document interactively
                redaction = parse_redaction_doc(doc_id)
                if not redaction:
                    raise ValueError("Failed to parse redaction document")
                
                # Get preview of redaction for headline selection context
                redaction_preview = ' '.join(redaction.split()[:30])
                
                # Now select headline interactively from cached headlines
                headline = select_headline_interactively(headlines_cache, row, redaction_preview)
                
                # Select cutline interactively from cached cutlines
                cutlines = select_cutline_interactively(cutlines_cache, headline)
                
                # Create sections dictionary for compatibility with existing code
                sections = {
                    'Headline': headline,
                    'Redaction': redaction,
                    'Cutlines': cutlines,
                    'Featured image': ''
                }
                
                # Update post info with headline
                post_info['headline'] = headline

                # Handle featured image - now with fallback mechanism
                featured_media_id = None
                image_caption = sections.get('Cutlines', '')
                
                # First attempt with the image URL from spreadsheet Column N
                if row.get('image_url'):
                    print(f"{BLUE}Attempting to use image URL from spreadsheet (Column N)...{ENDC}")
                    image_url = row['image_url']
                    featured_media_id = process_image_from_url(image_url, image_caption, doc_id)
                    
                    if featured_media_id:
                        post_info['image_status']['has_image'] = True
                        post_info['image_status']['status'] = 'Uploaded successfully from spreadsheet URL'
                        post_info['image_status']['media_id'] = featured_media_id
                        print(f"{GREEN}Successfully uploaded image with ID: {featured_media_id}{ENDC}")
                    else:
                        print(f"{YELLOW}Initial image upload from spreadsheet URL failed. Offering alternatives...{ENDC}")
                        
                        # Enable manual fallback for image upload
                        featured_media_id = handle_image_fallback(image_caption, doc_id)
                        
                        if featured_media_id:
                            post_info['image_status']['has_image'] = True
                            post_info['image_status']['status'] = 'Uploaded successfully via fallback method'
                            post_info['image_status']['media_id'] = featured_media_id
                        else:
                            post_info['image_status']['status'] = 'All image upload attempts failed'
                else:
                    print(f"{YELLOW}No image URL found in Column N. Offering alternatives...{ENDC}")
                    
                    # Enable manual fallback for image upload since no URL is available
                    featured_media_id = handle_image_fallback(image_caption, doc_id)
                    
                    if featured_media_id:
                        post_info['image_status']['has_image'] = True
                        post_info['image_status']['status'] = 'Uploaded successfully via manual input'
                        post_info['image_status']['media_id'] = featured_media_id

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

                # Determine the image source for display
                image_source = "Column N from spreadsheet"
                if post_info['image_status']['status'] == 'Uploaded successfully via fallback method':
                    image_source = "Alternative URL (fallback)"
                elif post_info['image_status']['status'] == 'Uploaded successfully via manual input':
                    image_source = "Manual input (local file)"

                # Display post details and wait for keyboard input
                display_post_details(
                    sections, 
                    row, 
                    featured_media_id is not None,
                    image_source
                )
                
                # Wait for keyboard command
                while True:
                    print("\nWaiting for command...")
                    key = get_single_key()
                    
                    # Check for ESC key (ASCII 27)
                    if key == '\x1b':  # ESC key
                        print(f"{RED}Exiting program...{ENDC}")
                        return  # Exit the main function
                    
                    elif key in ['\r', '\n']:  # ENTER = Publish
                        print(f"{GREEN}Publishing post...{ENDC}")
                        # Create post content with pre-looked-up IDs
                        post_content = {
                            'Headline': sections['Headline'],
                            'Redaction': sections['Redaction'],
                            'featured_media_id': featured_media_id,
                            'author_id': author_id,
                            'category_ids': category_ids,
                            'status': 'publish'  # Set status to publish
                        }
                        
                        # Create and publish WordPress post
                        post_response = create_wordpress_post_with_details(post_content, status='publish')
                        if post_response['success']:
                            post_info['status'] = 'Published'
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
                            print(f"{GREEN}Post published successfully:{ENDC} {post_response['post_url']}")
                        else:
                            post_info['error_details'] = post_response['error']
                            post_info['status'] = 'Failed'
                            failed_posts.append(post_info)
                            print(f"{RED}Failed to publish post: {post_response['error']}{ENDC}")
                        break
                        
                    elif key in ['\b', '\x08', '\x7f']:  # BACKSPACE = Create as Draft
                        print(f"{YELLOW}Creating post as draft...{ENDC}")
                        # Create post content with pre-looked-up IDs
                        post_content = {
                            'Headline': sections['Headline'],
                            'Redaction': sections['Redaction'],
                            'featured_media_id': featured_media_id,
                            'author_id': author_id,
                            'category_ids': category_ids,
                            'status': 'draft'  # Set status to draft
                        }
                        
                        # Create WordPress post as draft
                        post_response = create_wordpress_post_with_details(post_content, status='draft')
                        if post_response['success']:
                            post_info['status'] = 'Draft'
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
                            
                            # No need to update spreadsheet for drafts
                            post_info['sheet_update_status'] = 'Not updated (draft)'
                            
                            successful_posts.append(post_info)
                            print(f"{YELLOW}Post saved as draft:{ENDC} {post_response['post_url']}")
                        else:
                            post_info['error_details'] = post_response['error']
                            post_info['status'] = 'Failed'
                            failed_posts.append(post_info)
                            print(f"{RED}Failed to create draft: {post_response['error']}{ENDC}")
                        break
                        
                    elif key == ' ':  # SPACE = Skip
                        print(f"{BLUE}Skipping this post...{ENDC}")
                        post_info['status'] = 'Skipped'
                        skipped_posts.append(post_info)
                        break
                        
                    else:
                        print(f"{ORANGE}Unknown command. Please use ENTER, BACKSPACE, SPACE, or ESC.{ENDC}")

            except Exception as e:
                error_message = str(e)
                post_info['error_details'] = error_message
                post_info['status'] = 'Failed'
                failed_posts.append(post_info)
                print(f"{RED}Error processing row {row['row']}: {error_message}{ENDC}")
                print(f"{YELLOW}Press any key to continue to the next post...{ENDC}")
                get_single_key()
                continue

        # Print detailed summary grouped by section
        print(f"\n{BOLD}{BLUE}" + "="*70)
        print("POSTING SUMMARY BY SECTION")
        print("="*70 + f"{ENDC}")

        # Get all unique sections
        all_sections = set([post['section'] for post in successful_posts + failed_posts + skipped_posts])
        
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
                    print(f"\n{BOLD}Row {post['row']}: {post['headline']} ({post['status']}){ENDC}")
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
            
            # Filter skipped posts for this section
            section_skipped = [post for post in skipped_posts if post['section'] == section]
            if section_skipped:
                print(f"\n{BLUE}{BOLD}⏭️ SKIPPED POSTS{ENDC}")
                print("-"*50)
                for post in section_skipped:
                    print(f"Row {post['row']}: {post['headline']}")
            
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
            print(f"\n→ {BOLD}Section '{section}' summary:{ENDC} " +
                  f"{len([p for p in section_successful if p['status'] == 'Published'])} published, " +
                  f"{len([p for p in section_successful if p['status'] == 'Draft'])} draft, " +
                  f"{len(section_skipped)} skipped, " +
                  f"{len(section_failed)} failed")

        # Overall summary with percentages
        total_posts = len(successful_posts) + len(failed_posts) + len(skipped_posts)
        published_posts = len([p for p in successful_posts if p['status'] == 'Published'])
        draft_posts = len([p for p in successful_posts if p['status'] == 'Draft'])
        
        print(f"\n{BOLD}{BLUE}" + "="*70)
        print("OVERALL SUMMARY")
        print("="*70 + f"{ENDC}")
        print(f"{BOLD}Total sections:{ENDC} {len(all_sections)}")
        print(f"{BOLD}Total posts processed:{ENDC} {total_posts}")
        published_pct = (published_posts/total_posts*100 if total_posts > 0 else 0)
        draft_pct = (draft_posts/total_posts*100 if total_posts > 0 else 0)
        skipped_pct = (len(skipped_posts)/total_posts*100 if total_posts > 0 else 0)
        failed_pct = (len(failed_posts)/total_posts*100 if total_posts > 0 else 0)
        
        print(f"{BOLD}Published:{ENDC} {published_posts} ({published_pct:.1f}%)")
        print(f"{BOLD}Draft:{ENDC} {draft_posts} ({draft_pct:.1f}%)")
        print(f"{BOLD}Skipped:{ENDC} {len(skipped_posts)} ({skipped_pct:.1f}%)")
        print(f"{BOLD}Failed:{ENDC} {len(failed_posts)} ({failed_pct:.1f}%)")
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