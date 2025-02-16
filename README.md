# Google Drive to WordPress Automation

### Overview
This Python script automates the process of publishing content from Google Drive documents to WordPress. It monitors a Google Sheet for ready-to-publish articles, processes their content from Google Docs, handles featured images, and publishes them to WordPress while managing metadata like authors and categories.

### Prerequisites

- Python 3.x
- Google Cloud Project with enabled APIs:
  - Google Sheets API
  - Google Drive API
  - Google Docs API
- WordPress site with REST API access
- Required Python packages:
  ```
  google-api-python-client
  google-auth
  requests
  dotenv
  ```

Run the following command to install all packages:
```
pip install -r requirements.txt
```


### Configuration

#### Set up environmental variables:
1. Create '.env' file.
2. Add the following variables

```
GOOGLE_CREDENTIALS_FILE = /path/to/credentials.json  # Google service account credentials
WP_URL = https://your-wordpress-site.com/wp-json    # WordPress API endpoint
WP_USER = your_username                             # WordPress username
WP_PASSWORD = your_password                         # WordPress application password
```

### Google Sheet Structure

The script expects a Google Sheet with the following columns:
- Column B: "Ready To Post" (checkbox)
- Column C: "Online" (checkbox)
- Column D: "Story" (Google Doc URL)
- Column M: "Photographer" (optional)

### Google Doc Structure

Each Google Doc should contain the following sections:
```
Headline:
Featured image:
Cutlines:
Redaction:
Author(s):
Categories:
```

### Core Features

1. **Content Eligibility Detection**
   - Identifies articles ready for publishing
   - Checks "Ready To Post" status
   - Verifies article hasn't been published

2. **Content Extraction**
   - Parses Google Doc sections
   - Handles formatted text and hyperlinks
   - Processes multiple authors
   - Extracts category information

3. **Image Processing**
   - Downloads featured images from Google Drive
   - Uploads images to WordPress media library
   - Handles image captions and metadata

4. **WordPress Integration**
   - Creates formatted posts
   - Sets featured images
   - Assigns authors and categories
   - Publishes content with proper formatting

5. **Tracking and Reporting**
   - Updates Google Sheet status
   - Provides detailed success/failure reporting
   - Logs processing steps and errors

### Key Functions

#### `get_eligible_rows(sheet_id)`
Retrieves rows from Google Sheet that are ready for publishing.
- Parameters:
  - `sheet_id`: ID of the Google Sheet
- Returns: List of eligible row data

#### `parse_google_doc(doc_id)`
Extracts content sections from Google Doc.
- Parameters:
  - `doc_id`: ID of the Google Doc
- Returns: Dictionary of document sections

#### `create_wordpress_post(content_data)`
Creates a new WordPress post with formatted content.
- Parameters:
  - `content_data`: Dictionary containing post content and metadata
- Returns: Boolean indicating success/failure

#### `upload_image_to_wordpress(image_data, caption, filename)`
Uploads images to WordPress media library.
- Parameters:
  - `image_data`: Binary image data
  - `caption`: Image caption text
  - `filename`: Name for the uploaded file
- Returns: Media ID if successful, None if failed

### Error Handling

The script includes comprehensive error handling for:
- Network connectivity issues
- API rate limits
- Invalid content formatting
- Missing permissions
- File access errors
- WordPress API failures

### Output and Logging

The script provides detailed console output including:
- Processing status for each article
- Image upload confirmations
- Author and category assignments
- Success/failure summaries
- Error details for failed posts

### Limitations

1. Image Support
   - Only supports direct Google Drive image links
   - Processes single featured image per post

2. Content Formatting
   - Basic HTML formatting only
   - No support for complex layouts
   - Limited styling options

3. WordPress Integration
   - Requires WordPress REST API access
   - Application password authentication
   - Standard post type only
