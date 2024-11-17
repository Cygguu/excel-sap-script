/**
 * Created by: Jakub Zyzański
 * Version: 2.1
 *
 * This script updates specific cells in an Excel sheet based on 
 * SAP target values from column C.
 * It adjusts the content for yesterday's and today's target values, 
 * handling date formatting and specific conditions for certain entries
 * without affecting other formulas.
 *
 * The script processes values from column C, updating 
 * columns E, F, G, I, J, and H accordingly.
 * For rows where updates are applied, columns G and H are cleared.
 * Date formats are set to "yyyy/mm/dd" in columns F and J.
 *
 * Update from v1:
 * - Improved performance by reducing unnecessary operations and batch
 *   processing of data.
 * 
 * Update from v2:
 * - Added logic to skip updates on Sunday (Non Planifiée) for rows 
 *   containing "PR1" in column C.
 * - Enhanced functionality to clear cells in columns G and H for all 
 *   updated rows.
 * - Added clearing cells in columns K, L, and M for 
 *   entries "PS1" and "PBI".
 */

function main(workbook: ExcelScript.Workbook) {
    let sheet = workbook.getWorksheet("Compte-rendu Météo");

    // Arrays of target values for SAP with yesterday's and today's starting dates
    let targetValuesYesterday = [
        "P35", "P37", "P48", "P50", "PBO", "PBW", "PJ0", "PJ4",
        "PQ6", "PQ4", "PX0", "PX2", "PX4", "PY6", "PY8", "PC2",
        "PC0", "PD2", "PD0", "P01", "P83", "PK2", "PK0"
    ];
    let targetValuesToday = ["PA8", "P78"];

    // Get yesterday's and today's dates formatted as YYYY/MM/DD
    let yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    let formattedYesterday = yesterday.toISOString().slice(0, 10).replace(/-/g, '/');

    let today = new Date();
    let formattedToday = today.toISOString().slice(0, 10).replace(/-/g, '/');

    // Get the number of used rows in the sheet
    let lastRow = sheet.getUsedRange().getRowCount();

    // Retrieve values from columns to process in batch
    let rangeC = sheet.getRange("C1:C" + lastRow).getValues(); // Column C
    let rangeE = sheet.getRange("E1:E" + lastRow).getFormulas(); // Column E
    let rangeF = sheet.getRange("F1:F" + lastRow).getFormulas(); // Column F
    let rangeG = sheet.getRange("G1:G" + lastRow).getFormulas(); // Column G
    let rangeH = sheet.getRange("H1:H" + lastRow).getFormulas(); // Column H
    let rangeI = sheet.getRange("I1:I" + lastRow).getFormulas(); // Column I
    let rangeJ = sheet.getRange("J1:J" + lastRow).getFormulas(); // Column J

    // Get the current day of the week (0 = Sunday, 1 = Monday, ..., 6 = Saturday)
    let currentDay = new Date().getDay();

    // Iterate over rows and apply updates based on conditions
    for (let i = 0; i < lastRow; i++) {
        let cellValue = rangeC[i][0].toString(); // Get value from column C, row i

        if (cellValue === "PR1") {   // PR1 exception Non Planifiée on Sunday
            if (currentDay !== 0) { // Update on every day except Sunday
                rangeE[i][0] = "OK";
                rangeF[i][0] = formattedToday;
                rangeI[i][0] = 0;
                rangeJ[i][0] = formattedToday;
                rangeG[i][0] = ""; // Clear column G
                rangeH[i][0] = ""; // Clear column H
            }
        } else if (targetValuesYesterday.includes(cellValue)) {
            rangeE[i][0] = "OK";
            rangeF[i][0] = formattedYesterday;
            rangeI[i][0] = 0;
            rangeJ[i][0] = formattedYesterday;
            rangeG[i][0] = ""; // Clear column G
            rangeH[i][0] = ""; // Clear column H
        } else if (cellValue === "PA8") { // PA8 exception
            rangeE[i][0] = "OK";
            rangeF[i][0] = formattedToday;
            rangeI[i][0] = 9999;
            rangeJ[i][0] = formattedToday;
            rangeG[i][0] = "4:59"; // Set start time for PA8 "4:59"
            rangeH[i][0] = ""; // Clear column H
        } else if (targetValuesToday.includes(cellValue)) {
            rangeE[i][0] = "OK";
            rangeF[i][0] = formattedToday;
            rangeI[i][0] = 0;
            rangeJ[i][0] = formattedToday;
            rangeG[i][0] = ""; // Clear column G
            rangeH[i][0] = ""; // Clear column H
        } else if (cellValue === "P85") { // P85 exception
            rangeE[i][0] = "OK";
            rangeF[i][0] = formattedYesterday;
            rangeI[i][0] = 0;
            rangeJ[i][0] = formattedYesterday;
            rangeG[i][0] = ""; // Clear column G
            rangeH[i][0] = ""; // Clear column H
            // Set number format to "General"
            sheet.getRange("I" + (i + 1)).setNumberFormat("General"); 
            // Exception for PS1 and PBI
        } else if (cellValue === "PS1" || cellValue === "PBI") { 
            rangeE[i][0] = "OK";
            rangeF[i][0] = formattedYesterday;
            rangeI[i][0] = 0;
            rangeJ[i][0] = formattedYesterday;
            rangeG[i][0] = ""; // Clear column G
            rangeH[i][0] = ""; // Clear column H
            sheet.getRange("K" + (i + 1)).setValue(""); // Clear column K
            sheet.getRange("L" + (i + 1)).setValue(""); // Clear column L
            sheet.getRange("M" + (i + 1)).setValue(""); // Clear column M
        }
    }

    // Write back only the updated cells (values or formulas) in one operation
    sheet.getRange("E1:E" + lastRow).setFormulas(rangeE);
    sheet.getRange("F1:F" + lastRow).setFormulas(rangeF);
    sheet.getRange("G1:G" + lastRow).setFormulas(rangeG);
    sheet.getRange("H1:H" + lastRow).setFormulas(rangeH);
    sheet.getRange("I1:I" + lastRow).setFormulas(rangeI);
    sheet.getRange("J1:J" + lastRow).setFormulas(rangeJ);

    // Set the date format for the relevant columns
    sheet.getRange("F1:F" + lastRow).setNumberFormat("yyyy/mm/dd");
    sheet.getRange("J1:J" + lastRow).setNumberFormat("yyyy/mm/dd");
}
