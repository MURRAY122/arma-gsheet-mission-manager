## ArmaMissionGsheetManager

Monitor multiple directories and handles/manages a google spreadsheet with missions based on file names within the directory. This script was just a small project for me to work with gsheets but it might be of some use.

## Setup

Install pip packages via the requirments.txt file.

- pip install -r requirements.txt

Change the dirs_array to the required number of directories, Large number of dirs may cause higher usage and effect performace.

- dirs_array = [
  ["Example 1","Example1\\to\\path\\directory\\{0}".format(username)],  
   ["Example 2","Example2\\to\\path\\directory\\{0}".format(username)],
  ["Example 3","Example3\\to\\path\\directory\\{0}".format(username)]
  ]

Input DLC's in use by the servers (if any)

- map_dlcs = ["GM", "SOG", "WS"]

Create a google service account (https://support.google.com/a/answer/7378726?hl=en) and get the credentials.json file.

- CREDS_FILE_PATH = os.environ.get("CREDS_FILE_PATH") # Env variable
- CREDS_FILE_PATH = "Example/path/to/credentials.json" # Direct path

Create a google spreadsheet and obtain the ID from its URL, Leave the sheet Blank. You will also need to share the sheet with edit rights to your service account email (found in your credentials.json file).

- GSHEET_ID = os.environ.get("GSHEET_ID") # Env variable
- GSHEET_ID = "SheetsIdNumberAndLetters" # Direct id

After that you should be good to go.

## Mission File format

Files MUST be a .pbo type
File must follow the convention:
type max name version.map

for example:

- co 10 example v1.Altis
  If using DLC:
- GM ad 10 exmaple v2.Tanoa

A mission name CAN be the same as another within the same directory, so long as either the mission type or map is different.

For example:

- co 10 My Mission v1.Altis
- ad 10 My Mission v1.Altis
- co 10 My Mission v1.Tanoa
- co 10 My Mission v1.Malden
  Are all accetable within the same directory.

If a mission file is updated, This script will not delete the old version file from the directory but will update the spreadsheet with the updated version information.

- co 10 My Mission v1.Altis # Old Will be removed from sheet
- co 11 My Mission v2.Altis # Updated will be updated on sheet

## Last note

The top 3 rows of the spreadsheet will be skipped to allow for addition headers if required. However the default headers in row 2 must remain. Formatting can be done. Any rows the script adds to the sheet will copy the same format from the row above.
