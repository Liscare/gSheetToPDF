const gSpreadsheet = SpreadsheetApp.openById('ID of your GSheet')
const gPdfSheet = gSpreadsheet.getSheetByName('Name of a sheet')

var gSaveToRootFolder = true
var gChoices = ["?", "?"];
var gEmail = "liscare2@protonmail.com";

/**
 * Submit form handler.
 * Generate a PDF file from a gsheet according to the responses (email address and data choice).
 */
function onSubmitForm(event) {
    // Fetch form response
    let responseItems = event.response.getItemResponses()
    if (responseItems) {
        gChoices = responseItems[1].getResponse().split("-") // Can be Python - argparse
        gEmail = responseItems[0].getResponse() // Email address
    }
    // Insert gChoices into a google sheet to change data in the sheet (and then generate the PDF)
    let arr = [[gChoices[0], gChoices[1]]]
    // Cell R5 is for gChoices[0]
    // Cell S5 is for gChoices[1]
    // Other cells are linked to those 2 cells
    gPdfSheet.getRange('R5:S5').setValues(arr)
    SpreadsheetApp.flush()
    exportSelectedSheetAsPdf()
}

function exportSelectedSheetAsPdf() {
    let blob = getAsBlob(gSpreadsheet.getUrl(), gPdfSheet)
    sendEmailWithFile(blob, gPdfSheet.getName(), gSpreadsheet)
}

function sendEmailWithFile(blob, fileName, spreadsheet) {
    blob = blob.setName(fileName)
    let folder = gSaveToRootFolder ? DriveApp : DriveApp.getFileById(spreadsheet.getId()).getParents().next()
    let pdfFile = folder.createFile(blob)

    MailApp.sendEmail({
        to: gEmail,
        subject: "A subject",
        htmlBody: "Hi, I send you a PDf file from a Google Sheet.",
        attachments: [pdfFile.getAs(MimeType.PDF)]
    });
}

function getAsBlob(url, sheet, range) {
    let rangeParam = ''
    let sheetParam = ''
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
    let exportUrl = url.replace(/\/edit.*$/, '')
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

    let response
    let i = 0
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
