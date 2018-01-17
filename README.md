# Excel and Outlook Helpers

> Various modules for automating tasks in Excel and Outlook desktop applications

## Tested On
1. Windows 7 Professional
2. Outlook 2016
3. Excel 2016
4. Python 3.5.2

## email\_attachment\_export.py

> Module for exporting attachments from Outlook desktop application. I found myself needing to automate the process of downloading email attachments so I could import data contained in attachments into a SQL database. I created this module to help with that process.

### Requirements
* os
* win32com.client
* datetime

### Notes
* Use to download attachments from Outlook email messages
* If no path is specified, attachments will be downloaded to the Downloads folder
* Emails containing attachments can be searched for by the following:
 * keywords in the subject line
 * date received
 * day received relative to the current day
* Export filenames can be altered before export occurs

## excel\_refresh.py

> Module for refreshing data connections in Excel. I use this for refreshing data in Excel workbooks in place of manually opening the workbook, clicking on refresh all, and then saving. I also found that simply using the RefreshAll workbook method sometimes resulted in not all data being updated.
 
### Requirements
* os
* sys
* win32com.client
* time

### Notes
* If not using Excel 2016, change or remove the gencache.EnsureModule(
        '{00020813-0000-0000-C000-000000000046}', 0, 1, 9
        ) line
* To ensure all data is updated as expected, data is refreshed in a specified order based on connection and type of data

## email\_data\_from\_query.py

> Module for running a sql query, saving results to a file, attaching file to Outlook email tempate, and sending email. This seems to be a common task for me, so I created this module to automate the process.

### Requirements
* os
* win32com.client
* sql_stuff (one of my other modules)

### Notes
* After setting up sql_stuff, this module only requires paths to a .sql file, .msg email template, and file path for the to-be created data file