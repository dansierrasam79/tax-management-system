interface TaxpayerTax {
    taxpayerId: number;
    taxDue: number;
}

type CellValue = string | number | boolean;
function main(workbook: ExcelScript.Workbook) {
    const reportSheet =
        workbook.getWorksheet("Reports");

    const calculationsTable =
        workbook.getTable("tblTaxCalculations");

    // Load tax calculation data
    const calculationData: CellValue[][] =
        calculationsTable
            .getRangeBetweenHeaderAndTotal()
            .getValues();

    const taxpayers: TaxpayerTax[] = [];

    // Build an array of taxpayer objects
    for (const row of calculationData) {

        taxpayers.push({
            taxpayerId: Number(row[0]),
            taxDue: Number(row[5])
        });
    }

    // Sort highest tax first
    taxpayers.sort(
        (a, b) => b.taxDue - a.taxDue
    );

    // Report title
    reportSheet
        .getRange("J1")
        .setValue("Top Taxpayers");

    // Report headers
    reportSheet
        .getRange("J3:K3")
        .setValues([
            ["Taxpayer ID", "Tax Due"]
        ]);

    // Convert objects into Excel rows
    const reportData = taxpayers.map(
        taxpayer => [
            taxpayer.taxpayerId,
            taxpayer.taxDue
        ]
    );

    // Write report
    if (reportData.length > 0) {

        reportSheet
            .getRange(
                `J4:K${reportData.length + 3}`
            )
            .setValues(reportData);

    }

    // Format output
    reportSheet
        .getRange("J:K")
        .getFormat()
        .autofitColumns();
}
