type CellValue = string | number | boolean;

function main(workbook: ExcelScript.Workbook) {

    const reportSheet =
        workbook.getWorksheet("Reports");

    const taxpayerTable =
        workbook.getTable("tblTaxpayerInput");

    const calculationsTable =
        workbook.getTable("tblTaxCalculations");

    // Load table data
    const taxpayerData: CellValue[][] =
        taxpayerTable
            .getRangeBetweenHeaderAndTotal()
            .getValues();

    const calculationData: CellValue[][] =
        calculationsTable
            .getRangeBetweenHeaderAndTotal()
            .getValues();

    // TaxpayerID -> State
    const taxpayerStates =
        new Map<number, string>();

    // State -> Total Tax Due
    const stateTotals =
        new Map<string, number>();

    // Build TaxpayerID lookup
    for (const row of taxpayerData) {

        const taxpayerId = Number(row[0]);
        const state = String(row[3]);

        taxpayerStates.set(
            taxpayerId,
            state
        );
    }

    // Aggregate tax due by state
    for (const row of calculationData) {

        const taxpayerId = Number(row[0]);
        const taxDue = Number(row[5]);

        const state =
            taxpayerStates.get(taxpayerId);

        if (!state) {
            continue;
        }

        const currentTotal =
            stateTotals.get(state) ?? 0;

        stateTotals.set(
            state,
            currentTotal + taxDue
        );
    }

    // Report title
    reportSheet
        .getRange("P1")
        .setValue("State Analysis");

    // Report headers
    reportSheet
        .getRange("P3:Q3")
        .setValues([
            ["State", "Total Tax Due"]
        ]);

    const reportData: (string | number)[][] = [];

    // Convert Map into rows
    stateTotals.forEach(
        (totalTax, state) => {

            reportData.push([
                state,
                totalTax
            ]);

        }
    );

    // Sort highest tax due first
    reportData.sort(
        (a, b) =>
            Number(b[1]) - Number(a[1])
    );

    // Write results
    if (reportData.length > 0) {

        reportSheet
            .getRange(
                `P4:Q${reportData.length + 3}`
            )
            .setValues(reportData);

    }

    // Format worksheet
    reportSheet
        .getRange("P:Q")
        .getFormat()
        .autofitColumns();
}
