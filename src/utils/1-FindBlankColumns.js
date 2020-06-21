/* eslint-disable no-undef */
const findBlankColumns = async(name = "Sheet1") => {
    await Excel.run(async cxt => {
        const sheet = cxt.workbook.worksheets.getItem(name);
        sheet.getRange("A1").getEntireRow().insert(Excel.InsertShiftDirection.down);
        sheet.getRange("A1").values = [
            [`=IF(COUNTA(A2:A5000)=0,"Blank","Not blank")`]
        ];
        sheet.getRange("B1").getEntireRow().copyFrom("A1", Excel.RangeCopyType.formulas, true);
        await cxt.sync();
    })
}

export default findBlankColumns;