# DOCX to Document360

Converts a Word `.docx` file to clean HTML and uploads it to Document360.

## Setup

```powershell
python -m pip install -r requirements.txt
```

## Usage

Generate HTML only:

```powershell
python .\docx_to_document360.py --no-upload
```

Generate HTML and upload:

```powershell
$env:DOCUMENT360_API_TOKEN="your-token"
python .\docx_to_document360.py --publish
```

Or place credentials in a `.env` file in this folder:

```dotenv
DOCUMENT360_API_TOKEN=your-token
DOCUMENT360_USER_ID=your-user-id
DOCUMENT360_PROJECT_VERSION_ID=your-project-version-id
DOCUMENT360_CATEGORY_ID=your-category-id
DOCUMENT360_LANG_CODE=en
DOCUMENT360_PORTAL_URL=https://your-site.document360.io
```

Optional environment variables:

- `DOCUMENT360_BASE_URL` default: `https://apihub.document360.io`
- `DOCUMENT360_PROJECT_VERSION_ID`
- `DOCUMENT360_USER_ID`
- `DOCUMENT360_LANG_CODE`
- `DOCUMENT360_CATEGORY_ID`
- `DOCUMENT360_CATEGORY_NAME` default: `Imports`
- `DOCUMENT360_PUBLISH` default: `true`

If `DOCUMENT360_PROJECT_VERSION_ID`, `DOCUMENT360_USER_ID`, or `DOCUMENT360_LANG_CODE` are not set, the script auto-selects the main project version, the first available team account, and the default language returned by the API.
