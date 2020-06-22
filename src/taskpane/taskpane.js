/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

import FindBlankColumns from '../utils/1-FindBlankColumns'
import SortBlankColumns from '../utils/2-SortBlankColumns';
import RemoveBlankColumns from '../utils/3-RemoveBlankColumns';
import CreateTable from '../utils/4-CreateTable';

Office.onReady(info => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "flex";
        document.getElementById("run").onclick = run;
    }
});

export async function run() {
    try {
        await Excel.run(async context => {
            /**
             * Insert your Excel code here
             */
            // const range = context.workbook.getSelectedRange();

            // Read the range address
            // range.load("address");

            // Update the fill color
            // range.format.fill.color = "yellow";

            // ----- find blank columns -----

            await FindBlankColumns()
            const sheet = context.workbook.worksheets.getItem("Sheet1")

            // sheet.getRange("A1").getEntireRow().insert(Excel.InsertShiftDirection.down);
            // sheet.getRange("A1").values = [
            //     [`=IF(COUNTA(A2:A5000)=0,"Blank","Not blank")`]
            // ];
            // sheet.getRange("B1").getEntireRow().copyFrom("A1", Excel.RangeCopyType.formulas, true);

            // sheet.getRange("A:C").delete(Excel.DeleteShiftDirection.left);
            // const range = sheet.getRange("A1:BB5000");
            // const header = range.find("Not blank", {});
            // header.load("columnIndex");
            // await context.sync();
            // range.sort.apply(
            //     [{
            //         key: header.columnIndex,
            //         sortOn: Excel.SortOn.value
            //     }],
            //     false /*matchCase*/ ,
            //     false /*hasHeaders*/ ,
            //     Excel.SortOrientation.columns
            // );

            await SortBlankColumns()

            // const range1 = sheet.getRange("A1").getEntireRow();
            // const blankColumns = range1.find("Blank", {
            //     completeMatch: true
            // });
            // blankColumns.load("address");
            // await context.sync();
            // sheet.getRange(`${blankColumns.address[blankColumns.address.length - 2]}:CC`).clear()
            // sheet.getRange("A1").getEntireRow().delete(Excel.DeleteShiftDirection.up);

            await RemoveBlankColumns()



            // sheet.getRange("A2").getEntireRow().delete(Excel.DeleteShiftDirection.up)
            // sheet.getRange("A2").getEntireRow().delete(Excel.DeleteShiftDirection.up)
            // const lastColumn = sheet.getRange("A1").getEntireRow().find("Open Balance", {
            //     matchCase: true,
            //     completeMatch: true,
            //     searchDirection: "Forward"
            // });

            // const heightRange = sheet.getRange("H1:H1000");
            // const tableHeight = context.workbook.functions.countA(heightRange);

            // lastColumn.load("address");
            // tableHeight.load("value");
            // await context.sync();
            // console.log(lastColumn.address)
            // console.log(tableHeight.value)
            // const tableRange = `A1:${lastColumn.address[lastColumn.address.length - 2]}${tableHeight.value}`;
            // const table = context.workbook.tables.add(tableRange, true);
            // table.name = "POTable";

            // const tableHeader = table.getHeaderRowRange()



            // const dateColumn = tableHeader.find("Date", {}).getEntireColumn().delete(Excel.DeleteShiftDirection.left)
            // const numColumn = tableHeader.find("Num", {}).getEntireColumn().delete(Excel.DeleteShiftDirection.left)
            // const memoColumn = tableHeader.find("Memo", {}).getEntireColumn().delete(Excel.DeleteShiftDirection.left)
            // const sourceNameColumn = tableHeader.find("Source Name", {}).getEntireColumn().delete(Excel.DeleteShiftDirection.left)
            // const delivDateColumn = tableHeader.find("Deliv Date", {}).getEntireColumn().delete(Excel.DeleteShiftDirection.left)
            // const rcvColumn = tableHeader.find("Rcv'd", {}).getEntireColumn().delete(Excel.DeleteShiftDirection.left)
            // const backOrderColumn = tableHeader.find("Backordered", {}).getEntireColumn().delete(Excel.DeleteShiftDirection.left)
            // const amountColumn = tableHeader.find("Amount", {}).getEntireColumn().delete(Excel.DeleteShiftDirection.left)
            // const balanceColumn = tableHeader.find("Open Balance", {}).getEntireColumn().delete(Excel.DeleteShiftDirection.left)
            // const nameColumn = tableHeader.find("Name", {
            //     matchCase: true,
            //     completeMatch: true,
            //     searchDirection: "Forward"
            // }).getEntireColumn().delete(Excel.DeleteShiftDirection.left)

            // const a2 = sheet.getRange("A2")
            // a2.load('values')

            // await context.sync();
            // a2.values = [
            //     [`${a2.values[0][0].split(" (")[0]}`]
            // ]

            // console.log(`The range address was ${range.address}.`);

            await CreateTable()
        });
    } catch (error) {
        console.error(error);
    }
}