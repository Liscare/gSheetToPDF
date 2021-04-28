# gSheetToPDF
Send a PDF (from data in Google Sheet) by email according a choice in Google Form.
Many thanks to Jason Huang and his article [How to Print Google Sheet to PDF Using Apps Script](https://xfanatical.com/blog/print-google-sheet-as-pdf-using-apps-script/)
and also [Amit Agarwal](Amit Agarwal) and his article [How to Add Options in Google Forms Questions from Google Sheets](https://www.labnol.org/code/google-forms-choices-from-sheets-200630) for the populate form part.

## How it works
### Google Forms
A gForms with 2 questions: email address and a checkbox
The checkbox is a list of all data available. Can be a list of employers, for example.

Documentation of Google Sheet API [here](https://developers.google.com/apps-script/reference/forms)

### Google Sheet
A gSheet with 2 sheets:
- One with all data 
- One where the PDF will be generated using data from the other sheet

The data sheet is a list of products identified by two IDs (name and brand), one line for one product.
My PDF sheet refers to 2 cells (R5 and S5) via [`vlookup` function](https://support.google.com/docs/answer/3093318?hl=en). 
When the script will change those cells the PDF sheet will change (due to reference) and will display all data 
(from the data sheet) for the selected product.

Documentation of Google Sheet API [here](https://developers.google.com/sheets/api/quickstart/apps-script)

### App Script

After the form submission, the script in `main.js` will :
- Fetch responses from the gForms
- Actualize the sheet with those responses (preparation for the PDF)
- Create the PDF file from the sheet
- Send this PDF by email

You need to copy the content of `main.js` into a file in App Script from your Google Forms (to be linked), 
see [documentation](https://developers.google.com/apps-script/quickstart/forms).

Before sending the PDF by email, this file is stored into Google Drive for a moment.
