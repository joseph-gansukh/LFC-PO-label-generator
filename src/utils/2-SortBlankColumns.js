/* eslint-disable no-undef */
const sortBlankColumns = async(criteria = "Blank") => {
    await Excel.run(async cxt => {
        const sheet = cxt.workbook.worksheets.getActiveWorksheet();
        sheet.getRange("A:C").delete(Excel.DeleteShiftDirection.left);
        const range = sheet.getRange("A1:BB5000");
        const header = range.find(criteria, {});
        header.load("columnIndex");
        await cxt.sync();
        range.sort.apply(
            [{
                key: header.columnIndex,
                sortOn: Excel.SortOn.value
            }],
            false /*matchCase*/ ,
            false /*hasHeaders*/ ,
            Excel.SortOrientation.columns
        );
        await cxt.sync();
    })
}

export default sortBlankColumns;