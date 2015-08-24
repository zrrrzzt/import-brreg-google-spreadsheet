# import-brreg-google-spreadsheet
A Google Apps Script for importing records from brreg to a Google Spreadsheet.

Use this script to import records from The Central Coordinating Register for Legal Entities i Norway directly into your spreadsheet.
Visit Google Developers site to learn more about [Apps Script](https://developers.google.com/apps-script/)

## Installation
In your google apps spreadsheet open the "Tools" menu and select "Script editor".
From the blank page select "Blank project". 
Delete the prefilled code and copy everything from [brreg.gs](brreg.gs) into the editor.
Save the file and you are ready for action

## Usage
The script creates a new menu "BRREG".
The first time you run the script you will need to authorize it.
Click the menu and select "Import from brreg".
Enter "næringskode" in the first prompt and startdate in the next.

## About
The script uses the open API from [The Central Coordinating Register for Legal Entities i Norway](https://confluence.brreg.no/display/DBNPUB/API).
This API is a pilot and it might change in the near future.
The data is licensed through [NLOD](http://data.norge.no/nlod/no/1.0)(Norwegian License for Public Data) 

## License
MIT