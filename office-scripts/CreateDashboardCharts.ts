function main(workbook: ExcelScript.Workbook) {

  const dashboardSheet =
    workbook.getWorksheet("Dashboard");

  // Remove existing charts
  dashboardSheet
    .getCharts()
    .forEach(chart => chart.delete());

  createStateChart(workbook, dashboardSheet);
  createIncomeSourceChart(workbook, dashboardSheet);
  createBracketChart(workbook, dashboardSheet);
  createTopTaxpayerChart(workbook, dashboardSheet);
  createHighestDeductionsChart(workbook, dashboardSheet);
  createNetIncomeChart(workbook, dashboardSheet);
}

/**
 * Tax Due by State
 */
function createStateChart(
  workbook: ExcelScript.Workbook,
  dashboardSheet: ExcelScript.Worksheet
): void {

  const sourceRange =
    workbook
      .getTable("tblStateAnalysis")
      .getRange();

  const chart =
    dashboardSheet.addChart(
      ExcelScript.ChartType.columnClustered,
      sourceRange
    );

  chart.setTop(200);
  chart.setLeft(20);
  chart.setHeight(250);
  chart.setWidth(500);

  chart.getTitle().setText(
    "Tax Due by State"
  );
}

/**
 * Income by Source
 */
function createIncomeSourceChart(
  workbook: ExcelScript.Workbook,
  dashboardSheet: ExcelScript.Worksheet
): void {

  const sourceRange =
    workbook
      .getTable("tblIncomeSourceAnalysis")
      .getRange();

  const chart =
    dashboardSheet.addChart(
      ExcelScript.ChartType.pie,
      sourceRange
    );

  chart.setTop(200);
  chart.setLeft(600);
  chart.setHeight(250);
  chart.setWidth(400);

  chart.getTitle().setText(
    "Income by Source"
  );
}

/**
 * Taxpayer Count by Tax Bracket
 */
function createBracketChart(
  workbook: ExcelScript.Workbook,
  dashboardSheet: ExcelScript.Worksheet
): void {

  const sourceRange =
    workbook
      .getTable("tblBracketAnalysis")
      .getRange();

  const chart =
    dashboardSheet.addChart(
      ExcelScript.ChartType.barClustered,
      sourceRange
    );

  chart.setTop(500);
  chart.setLeft(20);
  chart.setHeight(250);
  chart.setWidth(500);

  chart.getTitle().setText(
    "Tax Bracket Distribution"
  );
}

/**
 * Top Taxpayers by Tax Due
 */
function createTopTaxpayerChart(
  workbook: ExcelScript.Workbook,
  dashboardSheet: ExcelScript.Worksheet
): void {

  const sourceRange =
    workbook
      .getTable("tblTopTaxpayers")
      .getRange();

  const chart =
    dashboardSheet.addChart(
      ExcelScript.ChartType.columnClustered,
      sourceRange
    );

  chart.setTop(500);
  chart.setLeft(600);
  chart.setHeight(250);
  chart.setWidth(450);

  chart.getTitle().setText(
    "Top Taxpayers"
  );
}

/**
 * Highest Deductions
 */
function createHighestDeductionsChart(
  workbook: ExcelScript.Workbook,
  dashboardSheet: ExcelScript.Worksheet
): void {

  const sourceRange =
    workbook
      .getTable("tblHighestDeductions")
      .getRange();

  const chart =
    dashboardSheet.addChart(
      ExcelScript.ChartType.columnClustered,
      sourceRange
    );

  chart.setTop(800);
  chart.setLeft(20);
  chart.setHeight(250);
  chart.setWidth(500);

  chart.getTitle().setText(
    "Highest Deductions"
  );
}

/**
 * Net Taxable Income by State
 */
function createNetIncomeChart(
  workbook: ExcelScript.Workbook,
  dashboardSheet: ExcelScript.Worksheet
): void {

  const sourceRange =
    workbook
      .getTable("tblNetIncomeAnalysis")
      .getRange();

  const chart =
    dashboardSheet.addChart(
      ExcelScript.ChartType.columnClustered,
      sourceRange
    );

  chart.setTop(800);
  chart.setLeft(600);
  chart.setHeight(250);
  chart.setWidth(450);

  chart.getTitle().setText(
    "Net Taxable Income by State"
  );
}
