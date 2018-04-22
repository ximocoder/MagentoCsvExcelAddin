using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;

namespace ExcelAdd4Magento
{
    public partial class HusseRibbon
    {
        private List<MagentoDictionary> magentoDictionary;


        private void HusseRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            magentoDictionary = new List<MagentoDictionary>();
        }

        private void btnExportToMagento_Click(object sender, RibbonControlEventArgs e)
        {
            ReadExcel();
            if (!CheckExcel())
            {
                MessageBox.Show("Cannot export to Magento csv, the file contains errors", "Error", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }
            FileDialog fileDialog = new SaveFileDialog();
            fileDialog.AddExtension = true;
            fileDialog.DefaultExt = "csv";
            fileDialog.Filter = "Magento translation csv files (*.csv)|*.csv|All files (*.*)|*.*";
            fileDialog.FilterIndex = 1;
            fileDialog.Title = "Magento csv File Dialog";

            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                SaveCsv(fileDialog.FileName);
            }

        }

        private bool CheckExcel()
        {
            bool isOk = magentoDictionary.Count != 0;
            string errorlines = "";
            foreach (var lineDict in magentoDictionary)
            {
                if (!firstAndLastCharAreOk(lineDict))
                {
                    isOk = false;
                    errorlines += lineDict.Line + ", ";
                }
            }
            if (!isOk)
                MessageBox.Show("Remove all the symbols \" before exporting to Magento csv, lines: " + errorlines, "Warning",
                    MessageBoxButtons.OK, MessageBoxIcon.Stop);
            return isOk;
        }

        private bool firstAndLastCharAreOk(MagentoDictionary lineDict)
        {
            if (lineDict.ColumnA[0] == '\"' || lineDict.ColumnA[lineDict.ColumnA.Length - 1] == '\"')
                return false;
            if (lineDict.ColumnB[0] == '\"' || lineDict.ColumnB[lineDict.ColumnB.Length - 1] == '\"')
                return false;
            return true;
        }
        
        private void SaveCsv(string fileDialogFileName)
        {
            using (var w = new StreamWriter(fileDialogFileName))
            {
                foreach (var lineDict in magentoDictionary)
                {
                    var line = string.Format("\"{0}\",\"{1}\"", lineDict.ColumnA, lineDict.ColumnB);
                    w.WriteLine(line);
                    w.Flush();
                }
            }
            MessageBox.Show("File saved in: " + fileDialogFileName);
        }

        private void ReadExcel()
        {
            var app = Globals.ThisAddIn.Application;
            magentoDictionary.Clear();

            Worksheet activeWorksheet = app.ActiveSheet;
            int rowCount = activeWorksheet.UsedRange.Rows.Count;
            Range firstColumn = activeWorksheet.Range["A1:A" + rowCount];
            Range secondColumn = activeWorksheet.Range["B1:B" + rowCount];

            for (int idxRow = 1; idxRow <= rowCount; idxRow++)
            {
                string colA = ((Range)activeWorksheet.Cells[idxRow, 1]).Value2;
                string colB = ((Range)activeWorksheet.Cells[idxRow, 2]).Value2;
                magentoDictionary.Add(new MagentoDictionary() { Line = idxRow, ColumnA = colA, ColumnB = colB });
            }
        }
    }

    internal class MagentoDictionary
    {
        public int Line;
        public string ColumnA;
        public string ColumnB;
    }
}
