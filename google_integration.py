import re
import os
import sys
import termios
import tty
from googleapiclient.discovery import build
from google.oauth2 import service_account
from googleapiclient.http import MediaIoBaseDownload
from googleapiclient.errors import HttpError

from constants import (
    sheets_service, docs_service, GREEN, YELLOW, BLUE, ENDC, BOLD
)

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

def parse_redaction_doc(doc_id):
    """
    Parse redaction document with interactive line selection.
    Shows first 9 lines and lets user select where redaction starts.
    """
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
    Format expected: 
    "Headlines
    SECTION:
    identifier: headline text"
    """
    try:
        # Fetch the Google Doc content with ALL tabs
        doc = docs_service.documents().get(
            documentId=doc_id, 
            includeTabsContent=True
        ).execute()
        
        # Check if we have tabs in the document
        if 'tabs' in doc and doc['tabs']:
            print(f"Found {len(doc['tabs'])} tabs in the document")
            
            # Extract all lines from each tab in the document
            all_tabs_content = []
            
            for tab_idx, tab in enumerate(doc['tabs']):
                tab_lines = []
                print(f"Processing tab {tab_idx + 1}: {tab.get('title', 'Unnamed tab')}")
                
                # Skip tabs without document content
                if 'documentTab' not in tab:
                    continue
                    
                # Get content from this tab
                tab_content = tab['documentTab'].get('body', {}).get('content', [])
                
                # Extract text from each paragraph in this tab
                for element in tab_content:
                    if 'paragraph' in element:
                        elements = element['paragraph']['elements']
                        text = ''.join([e.get('textRun', {}).get('content', '') for e in elements])
                        if text.strip():  # Only append non-empty lines
                            tab_lines.append(text.strip())
                
                if tab_lines:
                    all_tabs_content.append({
                        'tab_idx': tab_idx,
                        'title': tab.get('title', f"Tab {tab_idx+1}"),
                        'lines': tab_lines
                    })
            
            # Look for the Headlines tab/section in each tab
            headlines = []
            headlines_found = False
            insides_section_found = False
            
            for tab_content in all_tabs_content:
                # Skip the entire tab if it doesn't have relevant content
                if not any(line.lower() == "headlines" for line in tab_content['lines']) and \
                   not any(line.lower() == "insides" for line in tab_content['lines']):
                    continue
                
                current_category = "Uncategorized"
                in_insides_section = False
                past_insides_section = False
                in_headlines_section = False
                
                # Process each line in the tab
                for line in tab_content['lines']:
                    # Check for section markers
                    if line.lower() == "insides":
                        print(f"Found Insides section (examples - will be skipped)")
                        in_insides_section = True
                        insides_section_found = True
                        continue
                    
                    if line.lower() == "headlines":
                        print(f"Found Headlines section in tab '{tab_content['title']}'")
                        in_headlines_section = True
                        in_insides_section = False  # We're past the insides section
                        past_insides_section = True
                        headlines_found = True
                        continue
                        
                    # Only process headlines if we're not in the Insides section
                    # and either:
                    # 1. We've found the Headlines section, OR
                    # 2. We've passed the Insides section but haven't explicitly found a Headlines section
                    if not in_insides_section and (in_headlines_section or past_insides_section):
                        # Check if this is a new section marker (ends with colon)
                        if line.endswith(':'):
                            current_category = line.rstrip(':')
                            print(f"Found headline category: {current_category}")
                            continue
                        
                        # Skip empty lines or lines without a colon
                        if not line or not ':' in line:
                            continue
                            
                        # Parse headline format "identifier: headline text"
                        parts = line.split(':', 1)
                        if len(parts) == 2:
                            identifier = parts[0].strip()
                            headline_text = parts[1].strip()
                            
                            # Handle "SH:" in headline text
                            if "SH:" in headline_text:
                                parts = headline_text.split('SH:', 1)
                                headline_text = f"{parts[0].strip()}: {parts[1].strip()}"
                            
                            headlines.append({
                                'slug': identifier,
                                'headline': headline_text,
                                'category': current_category,
                                'original': line
                            })
                            print(f"Found headline: {identifier} - {headline_text}")
                
                # If we found headlines in this tab, don't check other tabs
                if headlines_found:
                    break
            
            # If no Headlines section found in any tab, try NEWS: markers
            if not headlines_found and not insides_section_found:
                for tab_content in all_tabs_content:
                    for i, line in enumerate(tab_content['lines']):
                        if line == "NEWS:":
                            print(f"Found NEWS: marker in tab '{tab_content['title']}', assuming this is part of Headlines section")
                            current_category = "NEWS"
                            
                            # Process subsequent lines as headlines
                            for next_line in tab_content['lines'][i+1:]:
                                if not ':' in next_line or next_line.endswith(':'):
                                    if next_line.endswith(':'):
                                        current_category = next_line.rstrip(':')
                                    continue
                                
                                parts = next_line.split(':', 1)
                                if len(parts) == 2:
                                    identifier = parts[0].strip()
                                    headline_text = parts[1].strip()
                                    
                                    headlines.append({
                                        'slug': identifier,
                                        'headline': headline_text,
                                        'category': current_category,
                                        'original': next_line
                                    })
                                    print(f"Found headline: {identifier} - {headline_text}")
                            
                            headlines_found = True
                            break
                    if headlines_found:
                        break
            
            if not headlines_found:
                print(f"{YELLOW}Could not find Headlines section in any tab{ENDC}")
                return []
        else:
            # Fallback to the old method (document without tabs)
            print("Document doesn't have tabs. Falling back to single document parser.")
            content = doc['body']['content']
            
            all_lines = []
            
            for element in content:
                if 'paragraph' in element:
                    elements = element['paragraph']['elements']
                    text = ''.join([e.get('textRun', {}).get('content', '') for e in elements])
                    if text.strip():  # Only append non-empty lines
                        all_lines.append(text.strip())
            
            # Look for the "Headlines" section anywhere in the document
            headlines_start = None
            headlines = []
            insides_section_found = False
            
            # First, check for "Insides" section
            for i, line in enumerate(all_lines):
                if line.lower() == "insides":
                    insides_section_found = True
                    print("Found Insides section (examples - will be skipped)")
                    break
            
            # Look for Headlines section
            for i, line in enumerate(all_lines):
                if line.lower() == "headlines":
                    headlines_start = i + 1
                    print(f"Found Headlines section at line {headlines_start}")
                    break
            
            # If Headlines found, parse content
            if headlines_start:
                current_category = "Uncategorized"
                
                for line in all_lines[headlines_start:]:
                    # Check for end markers
                    if line.lower() == "newscast" or line.lower() == "cutlines":
                        break
                        
                    # Check if this is a new section marker
                    if line.endswith(':'):
                        current_category = line.rstrip(':')
                        print(f"Found headline category: {current_category}")
                        continue
                    
                    # Skip empty lines or lines without a colon
                    if not line or not ':' in line:
                        continue
                        
                    # Parse headline format
                    parts = line.split(':', 1)
                    if len(parts) == 2:
                        identifier = parts[0].strip()
                        headline_text = parts[1].strip()
                        
                        # Handle "SH:" in headline text
                        if "SH:" in headline_text:
                            parts = headline_text.split('SH:', 1)
                            headline_text = f"{parts[0].strip()}: {parts[1].strip()}"
                        
                        headlines.append({
                            'slug': identifier,
                            'headline': headline_text,
                            'category': current_category,
                            'original': line
                        })
                        print(f"Found headline: {identifier} - {headline_text}")
            else:
                print(f"{YELLOW}Could not find Headlines section in document{ENDC}")
        
        print(f"Found {len(headlines)} potential headlines in document")
        return headlines
        
    except Exception as e:
        print(f"Error parsing headlines document: {str(e)}")
        import traceback
        traceback.print_exc()
        return []

def parse_cutlines_doc(doc_id):
    """
    Parse cutlines document and return a list of cutline options.
    Format expected: 
    "SECTION:
    *identifier: Cutline text PHOTO CREDIT: credit info
    identifier: Cutline text PHOTO CREDIT: credit info"
    """
    try:
        # Fetch the Google Doc content with ALL tabs
        doc = docs_service.documents().get(
            documentId=doc_id, 
            includeTabsContent=True
        ).execute()
        
        # Check if we have tabs in the document
        if 'tabs' in doc and doc['tabs']:
            print(f"Found {len(doc['tabs'])} tabs in the document")
            
            # Extract all lines from each tab in the document
            all_tabs_content = []
            
            for tab_idx, tab in enumerate(doc['tabs']):
                tab_lines = []
                print(f"Processing tab {tab_idx + 1}: {tab.get('title', 'Unnamed tab')}")
                
                # Skip tabs without document content
                if 'documentTab' not in tab:
                    continue
                    
                # Get content from this tab
                tab_content = tab['documentTab'].get('body', {}).get('content', [])
                
                # Extract text from each paragraph in this tab
                for element in tab_content:
                    if 'paragraph' in element:
                        elements = element['paragraph']['elements']
                        text = ''.join([e.get('textRun', {}).get('content', '') for e in elements])
                        if text.strip():  # Only append non-empty lines
                            tab_lines.append(text.strip())
                
                if tab_lines:
                    all_tabs_content.append({
                        'tab_idx': tab_idx,
                        'title': tab.get('title', f"Tab {tab_idx+1}"),
                        'lines': tab_lines
                    })
            
            # Look for the Cutlines tab/section in each tab
            cutlines = []
            cutlines_found = False
            
            for tab_content in all_tabs_content:
                # Look for "Cutlines" marker or indicators
                if any(line.lower() == "cutlines" for line in tab_content['lines']):
                    print(f"Found Cutlines section in tab '{tab_content['title']}'")
                    cutlines_found = True
                    
                    # Process lines in this tab
                    current_category = "Uncategorized"
                    
                    for line in tab_content['lines']:
                        # Skip the "Cutlines" marker line
                        if line.lower() == "cutlines":
                            continue
                            
                        # Check if this is a new section marker (ends with colon)
                        if line.endswith(':') and not ':' in line[:-1]:  # Only category markers end with : and don't have : before
                            current_category = line.rstrip(':')
                            print(f"Found cutline category: {current_category}")
                            continue
                            
                        # Skip empty lines
                        if not line:
                            continue
                        
                        # If line has a colon, it might be a cutline
                        if ':' in line:
                            # Remove leading asterisk if present
                            if line.startswith('*'):
                                line = line[1:].strip()
                            
                            # Split by first colon to get identifier and content
                            parts = line.split(':', 1)
                            if len(parts) == 2:
                                identifier = parts[0].strip()
                                cutline_text = parts[1].strip()
                                
                                # Extract photo credit if available
                                photo_credit = None
                                if "PHOTO CREDIT" in cutline_text:
                                    credit_parts = cutline_text.split("PHOTO CREDIT", 1)
                                    cutline_text = credit_parts[0].strip()
                                    if len(credit_parts) > 1:
                                        photo_credit = credit_parts[1].strip()
                                        if photo_credit.startswith(':'):
                                            photo_credit = photo_credit[1:].strip()
                                
                                # Also check for alternate "PHOTO CREDITS" format
                                elif "PHOTO CREDITS" in cutline_text:
                                    credit_parts = cutline_text.split("PHOTO CREDITS", 1)
                                    cutline_text = credit_parts[0].strip()
                                    if len(credit_parts) > 1:
                                        photo_credit = credit_parts[1].strip()
                                        if photo_credit.startswith(':'):
                                            photo_credit = photo_credit[1:].strip()
                                
                                cutlines.append({
                                    'slug': identifier,
                                    'cutline': cutline_text,
                                    'photo_credit': photo_credit,
                                    'category': current_category,
                                    'original': line
                                })
                                
                                print(f"Found cutline for {identifier}:")
                                print(f"  Text: {cutline_text}")
                                if photo_credit:
                                    print(f"  Credit: {photo_credit}")
                
                # If we found and processed cutlines in this tab, stop looking
                if cutlines_found:
                    break
                    
            # If no Cutlines section found in any tab, try with NEWS: marker directly
            if not cutlines_found:
                for tab_content in all_tabs_content:
                    for i, line in enumerate(tab_content['lines']):
                        if line == "NEWS:":
                            print(f"Found NEWS: marker in tab '{tab_content['title']}', assuming this is part of Cutlines section")
                            current_category = "NEWS"
                            cutlines_found = True
                            
                            # Process subsequent lines as potential cutlines
                            for j in range(i+1, len(tab_content['lines'])):
                                line = tab_content['lines'][j]
                                
                                # If this is a new section marker
                                if line.endswith(':') and not ':' in line[:-1]:
                                    current_category = line.rstrip(':')
                                    print(f"Found cutline category: {current_category}")
                                    continue
                                    
                                # Skip empty lines
                                if not line:
                                    continue
                                
                                # If line has a colon, it might be a cutline
                                if ':' in line:
                                    # Remove leading asterisk if present
                                    if line.startswith('*'):
                                        line = line[1:].strip()
                                    
                                    # Split by first colon to get identifier and content
                                    parts = line.split(':', 1)
                                    if len(parts) == 2:
                                        identifier = parts[0].strip()
                                        cutline_text = parts[1].strip()
                                        
                                        # Extract photo credit if available
                                        photo_credit = None
                                        if "PHOTO CREDIT" in cutline_text:
                                            credit_parts = cutline_text.split("PHOTO CREDIT", 1)
                                            cutline_text = credit_parts[0].strip()
                                            if len(credit_parts) > 1:
                                                photo_credit = credit_parts[1].strip()
                                                if photo_credit.startswith(':'):
                                                    photo_credit = photo_credit[1:].strip()
                                        
                                        # Also check for alternate "PHOTO CREDITS" format
                                        elif "PHOTO CREDITS" in cutline_text:
                                            credit_parts = cutline_text.split("PHOTO CREDITS", 1)
                                            cutline_text = credit_parts[0].strip()
                                            if len(credit_parts) > 1:
                                                photo_credit = credit_parts[1].strip()
                                                if photo_credit.startswith(':'):
                                                    photo_credit = photo_credit[1:].strip()
                                        
                                        cutlines.append({
                                            'slug': identifier,
                                            'cutline': cutline_text,
                                            'photo_credit': photo_credit,
                                            'category': current_category,
                                            'original': line
                                        })
                                        
                                        print(f"Found cutline for {identifier}:")
                                        print(f"  Text: {cutline_text}")
                                        if photo_credit:
                                            print(f"  Credit: {photo_credit}")
                            
                            break
                    
                    if cutlines_found:
                        break
            
            if not cutlines_found:
                print(f"{YELLOW}Could not find Cutlines section in any tab{ENDC}")
                return []
        else:
            # Fallback to the old method (document without tabs)
            print("Document doesn't have tabs. Falling back to single document parser.")
            content = doc['body']['content']
            
            all_lines = []
            
            for element in content:
                if 'paragraph' in element:
                    elements = element['paragraph']['elements']
                    text = ''.join([e.get('textRun', {}).get('content', '') for e in elements])
                    if text.strip():  # Only append non-empty lines
                        all_lines.append(text.strip())
            
            # Look for the "Cutlines" section anywhere in the document
            cutlines_start = None
            for i, line in enumerate(all_lines):
                if line.lower() == "cutlines":
                    cutlines_start = i + 1
                    print(f"Found Cutlines section at line {cutlines_start}")
                    break
                elif line == "NEWS:" and i > 0 and all_lines[i-1].lower() == "cutlines":
                    cutlines_start = i
                    print(f"Found NEWS: marker after Cutlines at line {cutlines_start}")
                    break
            
            # Process cutlines if section found
            cutlines = []
            if cutlines_start is not None:
                current_category = "Uncategorized"
                
                for line in all_lines[cutlines_start:]:
                    # Check if this is a new section marker
                    if line.endswith(':') and not ':' in line[:-1]:
                        current_category = line.rstrip(':')
                        print(f"Found cutline category: {current_category}")
                        continue
                        
                    # Skip empty lines
                    if not line:
                        continue
                    
                    # If line has a colon, it might be a cutline
                    if ':' in line:
                        # Remove leading asterisk if present
                        if line.startswith('*'):
                            line = line[1:].strip()
                        
                        # Split by first colon to get identifier and content
                        parts = line.split(':', 1)
                        if len(parts) == 2:
                            identifier = parts[0].strip()
                            cutline_text = parts[1].strip()
                            
                            # Extract photo credit if available
                            photo_credit = None
                            if "PHOTO CREDIT" in cutline_text:
                                credit_parts = cutline_text.split("PHOTO CREDIT", 1)
                                cutline_text = credit_parts[0].strip()
                                if len(credit_parts) > 1:
                                    photo_credit = credit_parts[1].strip()
                                    if photo_credit.startswith(':'):
                                        photo_credit = photo_credit[1:].strip()
                            
                            # Also check for alternate "PHOTO CREDITS" format
                            elif "PHOTO CREDITS" in cutline_text:
                                credit_parts = cutline_text.split("PHOTO CREDITS", 1)
                                cutline_text = credit_parts[0].strip()
                                if len(credit_parts) > 1:
                                    photo_credit = credit_parts[1].strip()
                                    if photo_credit.startswith(':'):
                                        photo_credit = photo_credit[1:].strip()
                            
                            cutlines.append({
                                'slug': identifier,
                                'cutline': cutline_text,
                                'photo_credit': photo_credit,
                                'category': current_category,
                                'original': line
                            })
                            
                            print(f"Found cutline for {identifier}:")
                            print(f"  Text: {cutline_text}")
                            if photo_credit:
                                print(f"  Credit: {photo_credit}")
            else:
                print(f"{YELLOW}Could not find Cutlines section in document{ENDC}")
        
        print(f"Found {len(cutlines)} potential cutlines in document")
        return cutlines
        
    except Exception as e:
        print(f"Error parsing cutlines document: {str(e)}")
        import traceback
        traceback.print_exc()
        return []

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
    """Extract the sheet ID from a Google Sheets URL."""
    match = re.search(r"/d/([a-zA-Z0-9_-]+)", sheet_url)
    if match:
        return match.group(1)
    else:
        raise ValueError("Invalid Google Sheets URL")
