interface ValidationIssue {
    severity: string;
    source: string;
    record: string;
    message: string;
}

type CellValue = string | number | boolean;

function main(workbook: ExcelScript.Workbook) {

    const reportsSheet = workbook.getWorksheet("Reports");

    const taxpayerTable = workbook.getTable("tblTaxpayerInput");
    const incomeTable = workbook.getTable("tblIncome");
    const deductionTable = workbook.getTable("tblDeductions");

    const taxpayerData =
        taxpayerTable.getRangeBetweenHeaderAndTotal().getValues();

    const incomeData =
        incomeTable.getRangeBetweenHeaderAndTotal().getValues();

    const deductionData =
        deductionTable.getRangeBetweenHeaderAndTotal().getValues();

    let issues: ValidationIssue[] = [];

    validateTaxpayers(taxpayerData, issues);
    validateIncome(incomeData, issues);
    validateDeductions(deductionData, issues);

    writeValidationReport(
        reportsSheet,
        issues
    );
}

/**
 * Validate taxpayer records.
 */
function validateTaxpayers(
    taxpayerData: CellValue[][],
    issues: ValidationIssue[]
): void {

    const taxpayerIds = new Set<number>();

    for (let i = 0; i < taxpayerData.length; i++) {

        const row = taxpayerData[i];

        const taxpayerId = Number(row[0]);
        const fullName = String(row[1] ?? "").trim();

        if (!taxpayerId) {
            issues.push({
                severity: "Error",
                source: "Taxpayer_Input",
                record: `Row ${i + 2}`,
                message: "Missing Taxpayer ID"
            });
        }

        if (fullName === "") {
            issues.push({
                severity: "Error",
                source: "Taxpayer_Input",
                record: `Row ${i + 2}`,
                message: "Missing Full Name"
            });
        }

        if (taxpayerId) {

            if (taxpayerIds.has(taxpayerId)) {
                issues.push({
                    severity: "Error",
                    source: "Taxpayer_Input",
                    record: taxpayerId.toString(),
                    message: "Duplicate Taxpayer ID"
                });
            }

            taxpayerIds.add(taxpayerId);
        }
    }
}

/**
 * Validate income records.
 */
function validateIncome(
    incomeData: CellValue[][],
    issues: ValidationIssue[]
): void {

    for (let i = 0; i < incomeData.length; i++) {

        const row = incomeData[i];

        const taxpayerId = Number(row[1]);
        const amount = Number(row[3]);

        if (!taxpayerId) {
            issues.push({
                severity: "Error",
                source: "Income_Details",
                record: `Row ${i + 2}`,
                message: "Missing Taxpayer ID"
            });
        }

        if (amount < 0) {
            issues.push({
                severity: "Error",
                source: "Income_Details",
                record: `Row ${i + 2}`,
                message: "Negative Income Amount"
            });
        }
    }
}

/**
 * Validate deduction records.
 */
function validateDeductions(
    deductionData: CellValue[][],
    issues: ValidationIssue[]
): void {

    for (let i = 0; i < deductionData.length; i++) {

        const row = deductionData[i];

        const taxpayerId = Number(row[1]);
        const amount = Number(row[3]);

        if (!taxpayerId) {
            issues.push({
                severity: "Error",
                source: "Deductions",
                record: `Row ${i + 2}`,
                message: "Missing Taxpayer ID"
            });
        }

        if (amount < 0) {
            issues.push({
                severity: "Error",
                source: "Deductions",
                record: `Row ${i + 2}`,
                message: "Negative Deduction Amount"
            });
        }
    }
}

/**
 * Write validation results to the Reports worksheet.
 */
function writeValidationReport(
    reportSheet: ExcelScript.Worksheet,
    issues: ValidationIssue[]
): void {

    reportSheet.getUsedRange()?.clear();

    reportSheet.getRange("A1").setValue("Validation Report");

    reportSheet.getRange("A3:D3").setValues([
        ["Severity", "Source", "Record", "Message"]
    ]);

    if (issues.length === 0) {

        reportSheet.getRange("A4").setValue(
            "No validation issues found."
        );

        return;
    }

    const reportData = issues.map(issue => [
        issue.severity,
        issue.source,
        issue.record,
        issue.message
    ]);

    reportSheet
        .getRange(`A4:D${issues.length + 3}`)
        .setValues(reportData);

    reportSheet.getUsedRange()?.getFormat().autofitColumns();
}
