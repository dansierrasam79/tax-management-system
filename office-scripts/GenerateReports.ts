type CellValue = string | number | boolean;

function main(workbook: ExcelScript.Workbook) {

    const reportSheet =
        workbook.getWorksheet("Reports");

    const taxpayerTable =
        workbook.getTable("tblTaxpayerInput");

    const calculationsTable =
        workbook.getTable("tblTaxCalculations");

    // Load calculation data
    const taxpayerData =
        taxpayerTable.getRangeBetweenHeaderAndTotal().getValues();

    const calculationData: CellValue[][] =
        calculationsTable.getRangeBetweenHeaderAndTotal().getValues();

    const totalTaxpayers =
        taxpayerData.length;

    let totalIncome = 0;
    let totalDeductions = 0;
    let totalTaxDue = 0;
    let totalTaxRate = 0;

    // Summarise calculation results
    for (const row of calculationData) {

        totalIncome += Number(row[1]);
        totalDeductions += Number(row[2]);
        totalTaxDue += Number(row[5]);
        totalTaxRate += Number(row[4]);
    }

    const averageTaxRate =
        calculationData.length > 0
            ? totalTaxRate / calculationData.length
            : 0;

    // Write report title
    reportSheet.getRange("F1")
        .setValue("Tax Summary");

    // Write headers
    reportSheet.getRange("F3:G3")
        .setValues([
            ["Metric", "Value"]
        ]);

    // Write summary metrics
    reportSheet.getRange("F4:G8")
        .setValues([
            ["Total Taxpayers", totalTaxpayers],
            ["Total Income", totalIncome],
            ["Total Deductions", totalDeductions],
            ["Total Tax Due", totalTaxDue],
            ["Average Tax Rate", averageTaxRate]
        ]);

    // Format columns
    reportSheet.getRange("F:G")
        .getFormat()
        .autofitColumns();
}
