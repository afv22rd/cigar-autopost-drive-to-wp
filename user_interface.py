import sys
import termios
import tty
import platform

from constants import (
    GREEN, YELLOW, RED, BLUE, ENDC, BOLD, ORANGE
)

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

def select_headline_interactively(headlines, row_info, redaction_preview):
    """
    Present headline options to user for interactive selection.
    Returns the selected headline text.
    """
    # If no headlines found
    if not headlines:
        print(f"{YELLOW}No headline options found. Please enter a headline manually:{ENDC}")
        return input("Headline: ").strip()
    
    print(f"\n{BLUE}{BOLD}Processing row {row_info['row']} (Section: {row_info['section']}){ENDC}")
    print(f"\n{BOLD}Redaction preview:{ENDC}")
    print(f"{redaction_preview[:150]}...")
    
    print(f"\n{BOLD}What is the headline of this post?{ENDC}")
    
    # Group headlines by category for easier selection
    headlines_by_category = {}
    for headline in headlines:
        category = headline.get('category', 'Uncategorized')
        if category not in headlines_by_category:
            headlines_by_category[category] = []
        headlines_by_category[category].append(headline)
    
    # Assign numbers 1, 2, 3, etc. to each headline
    choices = {}
    num_idx = 1
    
    # Print headlines grouped by category
    for category, category_headlines in headlines_by_category.items():
        print(f"\n{BOLD}{category}:{ENDC}")
        for headline in category_headlines:
            choices[str(num_idx)] = headline
            print(f"{BOLD}{num_idx}. {headline['slug']}: {headline['headline']}{ENDC}")
            num_idx += 1
    
    print(f"\n{YELLOW}Enter number (1-{len(choices)}) or type a custom headline:{ENDC}")
    
    user_input = input("> ").strip()
    
    # Check if the input is a valid number choice
    if user_input in choices:
        selected_headline = choices[user_input]
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
    # If no cutlines found or no image
    if not cutlines:
        print(f"{YELLOW}No cutline options found. Enter a cutline or press Enter to skip:{ENDC}")
        return input("Cutline: ").strip()
    
    print(f"\n{BOLD}What is the cutline for the featured image?{ENDC}")
    print(f"{BLUE}(For headline: {headline}){ENDC}")
    
    # Group cutlines by category for easier selection
    cutlines_by_category = {}
    for cutline in cutlines:
        category = cutline.get('category', 'Uncategorized')
        if category not in cutlines_by_category:
            cutlines_by_category[category] = []
        cutlines_by_category[category].append(cutline)
    
    # Assign numbers 1, 2, 3, etc. to each cutline
    choices = {}
    num_idx = 1
    
    # Print cutlines grouped by category
    for category, category_cutlines in cutlines_by_category.items():
        print(f"\n{BOLD}{category}:{ENDC}")
        for cutline in category_cutlines:
            choices[str(num_idx)] = cutline
            
            # Format display text with photo credit if available
            display_text = cutline['cutline']
            if cutline.get('photo_credit'):
                display_text += f" PHOTO CREDIT: {cutline['photo_credit']}"
                
            print(f"{BOLD}{num_idx}. {cutline['slug']}: {display_text}{ENDC}")
            num_idx += 1
    
    print(f"\n{YELLOW}Enter number (1-{len(choices)}) or type a custom cutline or press Enter to skip:{ENDC}")
    
    user_input = input("> ").strip()
    
    if not user_input:
        print(f"{YELLOW}Skipping cutline.{ENDC}")
        return ""
    
    # Check if the input is a valid number choice
    if user_input in choices:
        selected_cutline = choices[user_input]
        
        # Build complete cutline text including photo credit if available
        cutline_text = selected_cutline['cutline']
        if selected_cutline.get('photo_credit'):
            cutline_text += f" PHOTO CREDIT: {selected_cutline['photo_credit']}"
            
        print(f"{GREEN}Selected: {cutline_text}{ENDC}")
        return cutline_text
    else:
        # Treat input as custom cutline
        print(f"{GREEN}Using custom cutline: {user_input}{ENDC}")
        return user_input

def display_post_details(sections, row, featured_media_available=False, image_source="Column N from spreadsheet"):
    """Display post details for review in a formatted way."""
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
