using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;

namespace ExcelAdd4Magento
{
    public partial class HusseRibbon
    {
        private void HusseRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnExportToMagento_Click(object sender, RibbonControlEventArgs e)
        {
            FileDialog fileDialog = new SaveFileDialog();

            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                SaveExcel();

            }

        }

        private void SaveExcel()
        {
            var app = Globals.ThisAddIn.Application;

            Worksheet activeWorksheet = app.ActiveSheet;
            //Range firstRow = activeWorksheet.get_Range("A1");
            //firstRow.EntireRow.Insert(XlInsertShiftDirection.xlShiftDown);
            //Range newFirstRow = activeWorksheet.get_Range("A1");
            //newFirstRow.Value2 = "This text was added by using code";
            int rowCount = activeWorksheet.UsedRange.Rows.Count + 10;
            Range firstColumn = activeWorksheet.Range["A2:A" + rowCount];
            Range secondColumn = activeWorksheet.Range["B2:B" + rowCount];

            for (int idxRow = 1; idxRow <= rowCount; idxRow++)
            {
                var originalValue = ((Range) activeWorksheet.Cells[idxRow, 1]).Value2;

                ((Range)activeWorksheet.Cells[idxRow, 1]).Value2 = originalValue + "test1";
            }

            for (int idxRow = 1; idxRow <= rowCount; idxRow++)
            {
                var originalValue = ((Range)activeWorksheet.Cells[idxRow, 1]).Value2;

                ((Range)activeWorksheet.Cells[idxRow, 2]).Value2 = originalValue + "test2";
            }

        }
    }
}
