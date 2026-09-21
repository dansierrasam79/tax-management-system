function main(workbook: ExcelScript.Workbook) {

  const reportSheet =
    workbook.getWorksheet("Reports");

  const dashboardSheet =
    workbook.getWorksheet("Dashboard");

  // Dashboard Title
  dashboardSheet.getRange("B1:F1")
    .merge();

  dashboardSheet.getRange("B1")
    .setValue("Tax Management & Analysis System");

  dashboardSheet.getRange("B1")
    .getFormat()
    .getFont()
    .setBold(true);

  dashboardSheet.getRange("B1")
    .getFormat()
    .getFont()
    .setSize(18);

  // Read values from Reports
  const totalTaxpayers =
    reportSheet.getRange("G4").getValue();

  const totalIncome =
    reportSheet.getRange("G5").getValue();

  const totalDeductions =
    reportSheet.getRange("G6").getValue();

  const totalTaxDue =
    reportSheet.getRange("G7").getValue();

  const averageTaxRate =
    reportSheet.getRange("G8").getValue();

  // Create KPI Cards
  createCard(
    dashboardSheet,
    "B3:C5",
    "Total Taxpayers",
    totalTaxpayers,
    "#D9EAD3"
  );

  createCard(
    dashboardSheet,
    "E3:F5",
    "Total Income",
    totalIncome,
    "#CFE2F3"
  );

  createCard(
    dashboardSheet,
    "H3:I5",
    "Total Deductions",
    totalDeductions,
    "#FFF2CC"
  );

  createCard(
    dashboardSheet,
    "K3:L5",
    "Total Tax Due",
    totalTaxDue,
    "#F4CCCC"
  );

  createCard(
    dashboardSheet,
    "N3:O5",
    "Average Tax Rate",
    `${averageTaxRate}%`,
    "#D9D2E9"
  );
}

function createCard(
  sheet: ExcelScript.Worksheet,
  rangeAddress: string,
  title: string,
  value: string | number | boolean,
  color: string
): void {

  const cardRange =
    sheet.getRange(rangeAddress);

  cardRange.merge();

  cardRange.setValue(
    `${title}\n\n${value}`
  );

  const format =
    cardRange.getFormat();

  format.getFill().setColor(color);

  format.getFont().setBold(true);

  format.getFont().setSize(14);

  format.setHorizontalAlignment(
    ExcelScript.HorizontalAlignment.center
  );

  format.setVerticalAlignment(
    ExcelScript.VerticalAlignment.center
  );

  format.getBorders().forEach(border => {
    border.setStyle(
      ExcelScript.BorderLineStyle.continuous
    );
  });
}
