type CellValue = string | number | boolean;

function main(workbook: ExcelScript.Workbook) {

    const reportSheet =
        workbook.getWorksheet("Reports");

    const incomeTable =
        workbook.getTable("tblIncome");

    // Load income records
    const incomeData: CellValue[][] =
        incomeTable
            .getRangeBetweenHeaderAndTotal()
            .getValues();

    // Income Type -> Total Amount
    const incomeTotals =
        new Map<string, number>();

    // Aggregate income by type
    for (const row of incomeData) {

        const incomeType =
            String(row[2]);

        const amount =
            Number(row[3]);

        const currentTotal =
            incomeTotals.get(incomeType) ?? 0;

        incomeTotals.set(
            incomeType,
            currentTotal + amount
        );
    }

    // Report title
    reportSheet
        .getRange("S1")
        .setValue("Income Source Analysis");

    // Report headers
    reportSheet
        .getRange("S3:T3")
        .setValues([
            ["Income Type", "Total Income"]
        ]);

    const reportData: (string | number)[][] = [];

    // Convert Map to rows
    incomeTotals.forEach(
        (total, incomeType) => {

            reportData.push([
                incomeType,
                total
            ]);

        }
    );

    // Sort highest income source first
    reportData.sort(
        (a, b) =>
            Number(b[1]) - Number(a[1])
    );

    // Write results
    if (reportData.length > 0) {

        reportSheet
            .getRange(
                `S4:T${reportData.length + 3}`
            )
            .setValues(reportData);

    }

    // Format report
    reportSheet
        .getRange("S:T")
        .getFormat()
        .autofitColumns();
}
