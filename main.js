var saveToRootFolder = true
var email = "liscare2@protonmail.com";
var choices = ["?", "?"];

/**
 * Submit form handler
 * Generate a PDF file from a gsheet according to the responses (email address and data choice)
 */
function onSubmitForm(event) {
    // Fetch form response
    let responseItems = event.response.getItemResponses()
    if (responseItems) {
        choices = responseItems[1].getResponse().split("-") // Can be Python - argparse
        email = responseItems[0].getResponse() // Email address
    }
    // Insert choices into a google sheet to change data in the sheet (and then generate the PDF)
    let sheet = SpreadsheetApp.openById('ID of your GSheet');
    let arr = [[choices[0], choices[1]]]
    // Cell R5 is for choices[0]
    // Cell S5 is for choices[1]
    // Other cells are linked to those 2 cells
    sheet.getSheetByName('Name of a sheet').getRange('R5:S5').setValues(arr)
    SpreadsheetApp.flush() //
    exportCurrentSheetAsPDF()
}


function _exportBlob(blob, fileName, spreadsheet) {
    blob = blob.setName(fileName)
    var folder = saveToRootFolder ? DriveApp : DriveApp.getFileById(spreadsheet.getId()).getParents().next()
    var pdfFile = folder.createFile(blob)

    MailApp.sendEmail({
        to: email,
        subject: "A subject",
        htmlBody: "Hi, I send you a PDf file from a Google Sheet.",
        attachments: [pdfFile.getAs(MimeType.PDF)]
    });

}

function exportAsPDF() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
    var blob = _getAsBlob(spreadsheet.getUrl())
    _exportBlob(blob, spreadsheet.getName(), spreadsheet)
}

function _getAsBlob(url, sheet, range) {
    var rangeParam = ''
    var sheetParam = ''
    if (range) {
        rangeParam =
            '&r1=' + (range.getRow() - 1)
            + '&r2=' + range.getLastRow()
            + '&c1=' + (range.getColumn() - 1)
            + '&c2=' + range.getLastColumn()
    }
    if (sheet) {
        sheetParam = '&gid=' + sheet.getSheetId()
    }
    // A credit to https://gist.github.com/Spencer-Easton/78f9867a691e549c9c70
    // these parameters are reverse-engineered (not officially documented by Google)
    // they may break overtime.
    var exportUrl = url.replace(/\/edit.*$/, '')
        + '/export?exportFormat=pdf&format=pdf'
        + '&size=LETTER'
        + '&portrait=true'
        + '&fitw=true'
        + '&top_margin=0.75'
        + '&bottom_margin=0.75'
        + '&left_margin=0.7'
        + '&right_margin=0.7'
        + '&sheetnames=false&printtitle=false'
        + '&pagenum=UNDEFINED' // change it to CENTER to print page numbers
        + '&gridlines=true'
        + '&fzr=FALSE'
        + sheetParam
        + rangeParam

    var response
    var i = 0
    for (; i < 5; i += 1) {
        response = UrlFetchApp.fetch(exportUrl, {
            muteHttpExceptions: true,
            headers: {
                Authorization: 'Bearer ' +  ScriptApp.getOAuthToken(),
            },
        })
        if (response.getResponseCode() === 429) {
            // printing too fast, retrying
            Utilities.sleep(3000)
        } else {
            break
        }
    }

    if (i === 5) {
        throw new Error('Printing failed. Too many sheets to print.')
    }

    return response.getBlob()
}


function exportCurrentSheetAsPDF() {
    var spreadsheet = SpreadsheetApp.openById('ID of your GSheet')
    var currentSheet = spreadsheet.getSheetByName('Name of a sheet')

    let blob = _getAsBlob(spreadsheet.getUrl(), currentSheet)
    _exportBlob(blob, currentSheet.getName(), spreadsheet)
}
