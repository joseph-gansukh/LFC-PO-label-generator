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

        const dateColumn = tableHeader.find("Date", {}).getEntireColumn()
        if (dateColumn != null) dateColumn.delete(Excel.DeleteShiftDirection.left)

        const numColumn = tableHeader.find("Num", {}).getEntireColumn()
        if (numColumn != null) numColumn.delete(Excel.DeleteShiftDirection.left)

        const memoColumn = tableHeader.find("Memo", {}).getEntireColumn()
        if (memoColumn != null) memoColumn.delete(Excel.DeleteShiftDirection.left)

        const sourceNameColumn = tableHeader.find("Source Name", {}).getEntireColumn()
        if (sourceNameColumn != null) sourceNameColumn.delete(Excel.DeleteShiftDirection.left)

        const delivDateColumn = tableHeader.find("Deliv Date", {}).getEntireColumn()
        if (delivDateColumn != null) delivDateColumn.delete(Excel.DeleteShiftDirection.left)

        const rcvColumn = tableHeader.find("Rcv'd", {}).getEntireColumn()
        if (rcvColumn != null) rcvColumn.delete(Excel.DeleteShiftDirection.left)

        const backOrderColumn = tableHeader.find("Backordered", {}).getEntireColumn()
        if (backOrderColumn != null) backOrderColumn.delete(Excel.DeleteShiftDirection.left)

        const amountColumn = tableHeader.find("Amount", {}).getEntireColumn()
        if (amountColumn != null) amountColumn.delete(Excel.DeleteShiftDirection.left)

        const balanceColumn = tableHeader.find("Open Balance", {}).getEntireColumn()
        if (backOrderColumn != null) balanceColumn.delete(Excel.DeleteShiftDirection.left)

        const nameColumn = tableHeader.find("Name", {
            matchCase: true,
            completeMatch: true,
            searchDirection: "Forward"
        }).getEntireColumn()
        if (nameColumn != null) nameColumn.delete(Excel.DeleteShiftDirection.left)

        table.load('rows')
        table.load('columns')

        await context.sync()
            // console.log(table.rows.toJSON())

        let casketsObjs = []

        const tableColumns = table.columns.toJSON().items

        const casketNames = tableColumns[0].values.flat()
        const casketQty = tableColumns[1].values.flat()
        const barCodeNum = tableColumns[2].values.flat()

        for (let i = 1; i < (((tableHeight.value - 1) / 2 * 3) + 1); i = i + 3) {
            let casketNameArray = casketNames[i].split(" (")

            casketsObjs.push({
                casket: casketNameArray[0],
                qty: casketQty[i + 1],
                barCode: barCodeNum[i + 1]
            })

        }
        console.log(casketsObjs)





    })
}

export default createTable