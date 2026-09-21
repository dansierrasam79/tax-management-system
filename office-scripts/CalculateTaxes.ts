/********************************************************************
 * Type Alias
 *
 * CellValue is a custom type that represents the kinds of data
 * we expect to read from Excel cells.
 * Instead of repeatedly writing:
 * (string | number | boolean)
 *
 * we can create a shorter reusable name.
 ********************************************************************/
type CellValue = string | number | boolean;
/********************************************************************
 * MAIN ENTRY POINT
 *
 * Every Office Script must contain a main() function.
 *
 * Excel automatically starts execution here when the script runs.
 *
 * The workbook parameter represents the currently open workbook.
 ********************************************************************/
function main(workbook: ExcelScript.Workbook) {

  /****************************************************************
   * LOAD TABLE OBJECTS
   *
   * getTable() returns a Table object by name.
   *
   * Think of these as references to database tables stored
   * within the workbook.
   ***************************************************************/
  const incomeTable = workbook.getTable("tblIncome");
  const deductionTable = workbook.getTable("tblDeductions");
  const taxpayerTable = workbook.getTable("tblTaxpayerInput");
  const bracketTable = workbook.getTable("tblTaxBrackets");
  const calculationTable = workbook.getTable("tblTaxCalculations");

  /****************************************************************
   * CLEAR PREVIOUS CALCULATIONS
   *
   * Before calculating new results we remove old ones.
   *
   * getRangeBetweenHeaderAndTotal()
   * returns only the data rows.
   *
   * It excludes:
   * - Header Row
   * - Total Row
   ***************************************************************/
  //const dataBody = calculationTable.getRangeBetweenHeaderAndTotal();
  clearCalculationTable(calculationTable);
  /****************************************************************
   * Safety Check
   *
   * If no rows exist, there is nothing to clear.
   *
   * getRowCount() returns the number of data rows.
   ***************************************************************/
  //if (dataBody.getRowCount() > 0) {
    //dataBody.clear(
   //   ExcelScript.ClearApplyTo.contents
   // );
 // }

  /****************************************************************
   * READ TABLE DATA INTO MEMORY
   *
   * getValues() returns a 2-dimensional array.
   *
   * Example:
   *
   * [
   *   [1001, "Daniel"],
   *   [1002, "Sarah"]
   * ]
   *
   * This is much faster than repeatedly reading cells
   * from Excel during processing.
   ***************************************************************/
  const taxpayers =
    taxpayerTable.getRangeBetweenHeaderAndTotal().getValues();

  const incomeData: CellValue[][] =
    incomeTable.getRangeBetweenHeaderAndTotal().getValues();

  const deductionData: CellValue[][] =
    deductionTable.getRangeBetweenHeaderAndTotal().getValues();

  const bracketData: CellValue[][] =
    bracketTable.getRangeBetweenHeaderAndTotal().getValues();

  /****************************************************************
   * RESULTS ARRAY
   *
   * Instead of writing each calculation back to Excel one row
   * at a time, we'll build an array in memory.
   *
   * Once complete, all results are written in a single operation.
   *
   * This is a major performance best practice.
   ***************************************************************/
  let records: (string | number)[][] = [];

  /****************************************************************
   * PROCESS EACH TAXPAYER
   ***************************************************************/
  for (let taxpayer of taxpayers) {

    /************************************************************
     * taxpayer[0]
     *
     * Column 1 in tblTaxpayerInput = TaxpayerID
     *
     * Number() converts Excel values into numeric values.
     ***********************************************************/
    const taxpayerId = Number(taxpayer[0]);

    /************************************************************
     * Calculate income belonging to this taxpayer.
     ***********************************************************/
    const totalIncome: number =
      calculateTotalIncome(
        taxpayerId,
        incomeData
      );

    /************************************************************
     * Calculate deductions belonging to this taxpayer.
     ***********************************************************/
    const totalDeductions: number =
      calculateTotalDeductions(
        taxpayerId,
        deductionData
      );

    /************************************************************
     * TAXABLE INCOME
     *
     * Ensures taxable income never becomes negative.
     *
     * Example:
     * Income: 1,000
     * Deduction: 2,000
     *
     * Result:
     * 0
     ***********************************************************/
    const taxableIncome =
      Math.max(
        0,
        totalIncome - totalDeductions
      );

    /************************************************************
     * Determine which tax bracket applies.
     ***********************************************************/
    const taxRate: number =
      findTaxRate(
        taxableIncome,
        bracketData
      );

    /************************************************************
     * Tax Due Formula
     *
     * Example:
     *
     * Taxable Income: 100000
     * Rate: 20
     *
     * Tax Due:
     * 100000 × (20 / 100)
     ***********************************************************/
    const taxDue =
      taxableIncome * (taxRate / 100);

    /************************************************************
     * Store result in memory.
     *
     * The order must match the column order of
     * tblTaxCalculations.
     ***********************************************************/
    records.push([
      taxpayerId,
      totalIncome,
      totalDeductions,
      taxableIncome,
      taxRate,
      taxDue
    ]);
  }

  /****************************************************************
   * WRITE RESULTS TO EXCEL
   *
   * addRows(-1, data)
   *
   * -1 means append to the end of the table.
   ***************************************************************/
  if (records.length > 0) {
    calculationTable.addRows(
      -1,
      records
    );
  }

  /****************************************************************
   * CREATE AN AUDIT RECORD
   *
   * This allows us to track:
   * - When the script ran
   * - Which script ran
   * - Number of records processed
   * - Success or failure status
   ***************************************************************/
  logAudit(
    workbook,
    "CalculateTaxes",
    records.length,
    "Success",
    "Tax calculation completed successfully"
  );
}

/********************************************************************
 * CALCULATE TOTAL INCOME
 *
 * Loops through tblIncome and adds all income records
 * belonging to a single taxpayer.
 ********************************************************************/
function calculateTotalIncome(
  taxpayerId: number,
  incomeData: CellValue[][]
): number {

  let total = 0;

  for (let row of incomeData) {

    // Column 2 = TaxpayerID
    if (Number(row[1]) === taxpayerId) {

      // Column 4 = Amount
      total += Number(row[3]);
    }
  }

  return total;
}

/********************************************************************
 * CALCULATE TOTAL DEDUCTIONS
 *
 * Similar to calculateTotalIncome() but retrieves deduction amounts.
 ********************************************************************/
function calculateTotalDeductions(
  taxpayerId: number,
  deductionData: CellValue[][]
): number {

  let total = 0;

  for (let row of deductionData) {

    if (Number(row[1]) === taxpayerId) {
      total += Number(row[3]);
    }
  }

  return total;
}

/********************************************************************
 * FIND TAX RATE
 *
 * Searches tblTaxBrackets to determine which bracket
 * contains the taxpayer's taxable income.
 *
 * Example:
 *
 * Income = 850000
 *
 * Bracket:
 * 500001 - 1000000
 *
 * Rate Returned:
 * 20
 ********************************************************************/
function findTaxRate(
  taxableIncome: number,
  bracketData: CellValue[][]
): number {

  for (const row of bracketData) {

    const min = Number(row[1]);
    const max = Number(row[2]);
    const rate = Number(row[3]);

    if (
      taxableIncome >= min &&
      taxableIncome <= max
    ) {
      return rate;
    }
  }

  return 0;
}

function clearCalculationTable(
  calculationTable: ExcelScript.Table
): void {

  if (calculationTable.getRowCount() > 0) {

    calculationTable
      .getRangeBetweenHeaderAndTotal()
      .delete(ExcelScript.DeleteShiftDirection.up);
  }
}

/********************************************************************
 * AUDIT LOGGING
 * Adds a record to tblAuditLog so administrators can see
 * what happened and when.
 ********************************************************************/
function logAudit(
  workbook: ExcelScript.Workbook,
  scriptName: string,
  recordsProcessed: number,
  status: string,
  message: string
): void {

  const auditTable =
    workbook.getTable("tblAuditLog");

  auditTable.addRow(-1, [
    new Date().toISOString(),
    scriptName,
    recordsProcessed,
    status,
    message
  ]);
}
