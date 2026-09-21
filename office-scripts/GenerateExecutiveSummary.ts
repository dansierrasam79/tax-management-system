type CellValue = string | number | boolean;

function main(workbook: ExcelScript.Workbook) {

    const reportSheet =
        workbook.getWorksheet("Reports");

    const taxpayerTable =
        workbook.getTable("tblTaxpayerInput");

    const incomeTable =
        workbook.getTable("tblIncome");

    const calculationTable =
        workbook.getTable("tblTaxCalculations");

    const taxpayerData: CellValue[][] =
        taxpayerTable
            .getRangeBetweenHeaderAndTotal()
            .getValues();

    const incomeData: CellValue[][] =
        incomeTable
            .getRangeBetweenHeaderAndTotal()
            .getValues();

    const calculationData: CellValue[][] =
        calculationTable
            .getRangeBetweenHeaderAndTotal()
            .getValues();

    const totalTaxpayers =
        taxpayerData.length;

    let totalTaxRevenue = 0;
    let highestTaxDue = 0;
    let highestTaxpayer = 0;
    let totalTaxRate = 0;

    // Calculate tax metrics
    for (const row of calculationData) {

        const taxpayerId = Number(row[0]);
        const taxRate = Number(row[4]);
        const taxDue = Number(row[5]);

        totalTaxRevenue += taxDue;
        totalTaxRate += taxRate;

        if (taxDue > highestTaxDue) {
            highestTaxDue = taxDue;
            highestTaxpayer = taxpayerId;
        }
    }

    // Calculate average tax rate
    const averageTaxRate =
        calculationData.length > 0
            ? totalTaxRate / calculationData.length
            : 0;

    // State totals
    const taxpayerStates =
        new Map<number, string>();

    const stateTaxTotals =
        new Map<string, number>();

    for (const row of taxpayerData) {

        const taxpayerId = Number(row[0]);
        const state = String(row[3]);

        taxpayerStates.set(
            taxpayerId,
            state
        );
    }

    for (const row of calculationData) {

        const taxpayerId = Number(row[0]);
        const taxDue = Number(row[5]);

        const state =
            taxpayerStates.get(taxpayerId);

        if (!state) {
            continue;
        }

        const currentTotal =
            stateTaxTotals.get(state) ?? 0;

        stateTaxTotals.set(
            state,
            currentTotal + taxDue
        );
    }

    // Find highest tax state
    let highestTaxState = "";
    let highestStateTax = 0;

    stateTaxTotals.forEach(
        (totalTax, state) => {

            if (totalTax > highestStateTax) {
                highestStateTax = totalTax;
                highestTaxState = state;
            }

        }
    );

    // Income source totals
    const incomeSourceTotals =
        new Map<string, number>();

    for (const row of incomeData) {

        const incomeType =
            String(row[2]);

        const amount =
            Number(row[3]);

        const currentTotal =
            incomeSourceTotals.get(incomeType) ?? 0;

        incomeSourceTotals.set(
            incomeType,
            currentTotal + amount
        );
    }

    // Find largest income source
    let topIncomeSource = "";
    let topIncomeAmount = 0;

    incomeSourceTotals.forEach(
        (amount, source) => {

            if (amount > topIncomeAmount) {
                topIncomeAmount = amount;
                topIncomeSource = source;
            }

        }
    );

    // Report title
    reportSheet
        .getRange("AB1")
        .setValue("Executive Summary");

    // Headers
    reportSheet
        .getRange("AB3:AC3")
        .setValues([
            ["Metric", "Value"]
        ]);

    // KPI values
    reportSheet
        .getRange("AB4:AC10")
        .setValues([
            ["Total Taxpayers", totalTaxpayers],
            ["Total Tax Revenue", totalTaxRevenue],
            ["Highest Taxpayer", highestTaxpayer],
            ["Highest Tax Due", highestTaxDue],
            ["Highest Tax State", highestTaxState],
            ["Top Income Source", topIncomeSource],
            ["Average Tax Rate", averageTaxRate]
        ]);

    // Auto-fit columns
    reportSheet
        .getRange("AB:AC")
        .getFormat()
        .autofitColumns();
}
