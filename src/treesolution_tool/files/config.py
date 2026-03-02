# config.py

# Standarddateien (können im Menü überschrieben werden)
DEFAULT_USERS_FILE = "Benutzer.xlsx"
DEFAULT_USERS_SHEET = "Sheet1"
DEFAULT_KEYWORDS_FILE = "keywords_technische_accounts.txt"
DEFAULT_OUTPUT_FILE = "Upload.csv"

# Spalten im Benutzerexport
COL_ID = "id"
COL_USERNAME = "username"
COL_EMAIL = "email"
COL_FIRSTNAME = "firstname"
COL_LASTNAME = "lastname"
COL_INSTITUTION = "institution"
COL_DEPARTMENT = "department"
COL_AUTH = "auth"

# Fixwerte für Upload-Export
EXPORT_INSTITUTION_VALUE = "Sonic Suisse SA"
EXPORT_AUTH_VALUE = "iomadoidc"