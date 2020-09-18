/* eslint-disable no-undef */
/* eslint-disable no-unused-vars */

// Remove blank columns
const removeBlankColumns = async() => {
    await Excel.run(async cxt => {
        const sheet = cxt.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange("A1").getEntireRow();
        const blankColumns = range.find("Blank", {
            completeMatch: true
        });
        blankColumns.load("address");
        await cxt.sync();
        sheet.getRange(`${blankColumns.address[blankColumns.address.length - 2]}:CC`).clear()
        sheet.getRange("A1").getEntireRow().delete(Excel.DeleteShiftDirection.up);
        sheet.getRange("A1").values = [
            [
                "Casket Name"
            ]
        ]
        sheet.getRange("A2").getEntireRow().delete(Excel.DeleteShiftDirection.up)
        sheet.getRange("A2").getEntireRow().delete(Excel.DeleteShiftDirection.up)

        await cxt.sync();
    })
}

export default removeBlankColumns;
