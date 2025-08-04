import re
import requests
import os
import io
import time
from datetime import datetime
from base64 import b64encode
import random
import string

from constants import WP_URL, WP_USER, WP_PASSWORD, GREEN, YELLOW, RED, BLUE, ENDC, BOLD

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
