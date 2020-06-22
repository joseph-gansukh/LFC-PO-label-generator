/* eslint-disable no-undef */
const createTable = async() => {
    await Excel.run(async context => {
        const sheet = context.workbook.worksheets.getActiveWorksheet()
        const lastColumn = sheet.getRange("A1").getEntireRow().find("Open Balance", {
            matchCase: true,
            completeMatch: true,
            searchDirection: "Forward"
        });

        const heightRange = sheet.getRange("A1:A1000");
        const tableHeight = context.workbook.functions.countA(heightRange);

        lastColumn.load("address");
        tableHeight.load("value");
        await context.sync();
        console.log(lastColumn.address)
        console.log(tableHeight.value)
        const tableRange = `A1:${lastColumn.address[lastColumn.address.length - 2]}${((tableHeight.value - 1) / 2 * 3) + 1}`;
        const table = context.workbook.tables.add(tableRange, true);
        table.name = "POTable";

        sheet.getRange(`A${((tableHeight.value - 1) / 2 * 3) + 2}`).getEntireRow().delete(Excel.DeleteShiftDirection.up)
        sheet.getRange(`A${((tableHeight.value - 1) / 2 * 3) + 2}`).getEntireRow().delete(Excel.DeleteShiftDirection.up)
        sheet.getRange(`A${((tableHeight.value - 1) / 2 * 3) + 2}`).getEntireRow().delete(Excel.DeleteShiftDirection.up)

        const tableHeader = table.getHeaderRowRange()

        const dateColumn = tableHeader.find("Date", {}).getEntireColumn().delete(Excel.DeleteShiftDirection.left)
        const numColumn = tableHeader.find("Num", {}).getEntireColumn().delete(Excel.DeleteShiftDirection.left)
        const memoColumn = tableHeader.find("Memo", {}).getEntireColumn().delete(Excel.DeleteShiftDirection.left)
        const sourceNameColumn = tableHeader.find("Source Name", {}).getEntireColumn().delete(Excel.DeleteShiftDirection.left)
        const delivDateColumn = tableHeader.find("Deliv Date", {}).getEntireColumn().delete(Excel.DeleteShiftDirection.left)
        const rcvColumn = tableHeader.find("Rcv'd", {}).getEntireColumn().delete(Excel.DeleteShiftDirection.left)
        const backOrderColumn = tableHeader.find("Backordered", {}).getEntireColumn().delete(Excel.DeleteShiftDirection.left)
        const amountColumn = tableHeader.find("Amount", {}).getEntireColumn().delete(Excel.DeleteShiftDirection.left)
        const balanceColumn = tableHeader.find("Open Balance", {}).getEntireColumn().delete(Excel.DeleteShiftDirection.left)
        const nameColumn = tableHeader.find("Name", {
            matchCase: true,
            completeMatch: true,
            searchDirection: "Forward"
        }).getEntireColumn().delete(Excel.DeleteShiftDirection.left)


    })
}

export default createTable