import re
import os
import time
from constants import (
    GOOGLE_CREDENTIALS_FILE, WP_URL, WP_USER, WP_PASSWORD,
    GREEN, YELLOW, RED, BLUE, BOLD, ORANGE, ENDC
)

# Import functions from our modules
from google_integration import (
    get_eligible_rows, parse_redaction_doc, parse_headlines_doc,
    parse_cutlines_doc, get_sheet_id, update_online_status, get_single_key
)
from wordpress_integration import (
    get_or_create_author_id, get_category_ids, create_wordpress_post_with_details
)
from image_processing import (
    process_image_from_url, handle_image_fallback
)
from user_interface import (
    select_headline_interactively, select_cutline_interactively,
    display_post_details
)

def main(sheet_id):
    """Main function to process eligible posts with interactive keyboard controls."""
    successful_posts = []
    failed_posts = []
    skipped_posts = []  # New list to track skipped posts

    # Storage for headline and cutline options
    headlines_cache = []
    cutlines_cache = []
    
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
                # Handle URL with or without tab parameter
                headlines_doc_match = re.search(r'/document/d/([a-zA-Z0-9_-]+)', first_row['headlines_url'])
                if headlines_doc_match:
                    headlines_doc_id = headlines_doc_match.group(1)
                    print(f"Extracting headlines from document ID: {headlines_doc_id}")
                    print(f"Original URL: {first_row['headlines_url']}")
                    headlines_cache = parse_headlines_doc(headlines_doc_id)
                else:
                    print(f"{YELLOW}Invalid headlines document URL format.{ENDC}")
            else:
                print(f"{YELLOW}No headlines document URL found.{ENDC}")
                
            # Parse cutlines document (from first row's column Q)
            if first_row.get('cutlines_url'):
                print(f"{BLUE}Parsing cutlines document...{ENDC}")
                # Handle URL with or without tab parameter
                cutlines_doc_match = re.search(r'/document/d/([a-zA-Z0-9_-]+)', first_row['cutlines_url'])
                if cutlines_doc_match:
                    cutlines_doc_id = cutlines_doc_match.group(1)
                    print(f"Extracting cutlines from document ID: {cutlines_doc_id}")
                    print(f"Original URL: {first_row['cutlines_url']}")
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
                
                # Only ask for cutlines if there's an image URL in column N
                if row.get('image_url'):
                    # Select cutline interactively from cached cutlines
                    cutlines = select_cutline_interactively(cutlines_cache, headline)
                else:
                    # Skip cutline selection if no image URL
                    cutlines = ""
                    print(f"{YELLOW}No image URL found in Column N. Skipping cutline selection.{ENDC}")
                
                # Create sections dictionary for compatibility with existing code
                sections = {
                    'Headline': headline,
                    'Redaction': redaction,
                    'Cutlines': cutlines,
                    'Featured image': ''
                }
                
                # Update post info with headline
                post_info['headline'] = headline

                # Handle featured image - now with modified fallback mechanism
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
                        
                        # Enable manual fallback for image upload - only when the initial upload fails
                        featured_media_id = handle_image_fallback(image_caption, doc_id)
                        
                        if featured_media_id:
                            post_info['image_status']['has_image'] = True
                            post_info['image_status']['status'] = 'Uploaded successfully via fallback method'
                            post_info['image_status']['media_id'] = featured_media_id
                        else:
                            post_info['image_status']['status'] = 'All image upload attempts failed'
                else:
                    print(f"{YELLOW}No image URL found in Column N. Skipping image upload.{ENDC}")
                    # Don't offer fallback options when there's no image URL
                    post_info['image_status']['status'] = 'No image URL provided'

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
                image_source = "None"
                if row.get('image_url'):
                    if post_info['image_status']['status'] == 'Uploaded successfully from spreadsheet URL':
                        image_source = "Column N from spreadsheet"
                    elif post_info['image_status']['status'] == 'Uploaded successfully via fallback method':
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
    # Ask user for the spreadsheet URL
    sheet_url = input("Enter Google Sheets URL: ").strip()

    try:
        sheetid = get_sheet_id(sheet_url)
        print("Extracted Sheet ID:", sheetid)
    except ValueError as e:
        print(e)

    main(sheetid)