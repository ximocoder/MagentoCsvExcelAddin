using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

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
                try
                {
                    SaveCsv(fileDialog.FileName);
                }
                catch (Exception exception)
                {
                    MessageBox.Show(exception.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
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
            if (System.IO.File.Exists(fileDialogFileName))
            {
                MessageBox.Show("The file already exists. We will create a new one with .tmp extension");
                fileDialogFileName += ".tmp";
            }
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
            Application app = Globals.ThisAddIn.Application;
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
        private void CleanExcel()
        {
            Application app = Globals.ThisAddIn.Application;
            magentoDictionary.Clear();

            Worksheet activeWorksheet = app.ActiveSheet;
            activeWorksheet.Cells.Clear();
        }


        private void btnImportFromCsv_Click(object sender, RibbonControlEventArgs e)
        {
            FileDialog fileDialog = new OpenFileDialog();
            fileDialog.AddExtension = true;
            fileDialog.DefaultExt = "csv";
            fileDialog.Filter = "Magento translation csv files (*.csv)|*.csv|All files (*.*)|*.*";
            fileDialog.FilterIndex = 1;
            fileDialog.Title = "Magento csv File Dialog";

            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                CleanExcel();
                try
                {
                    ReadCsv(fileDialog.FileName);
                }
                catch (Exception exception)
                {
                    MessageBox.Show(exception.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void ReadCsv(string fileName)
        {
            List<string> listA = new List<string>();
            List<string> listB = new List<string>();
            Application app = Globals.ThisAddIn.Application;

            using (var reader = new StreamReader(fileName))
            {

                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(',');

                    if (values[0] != null)
                        listA.Add(values[0].Replace("\"", ""));
                    else
                        listA.Add("");
                    if (values[1] != null)
                        listB.Add(values[1].Replace("\"", ""));
                    else
                        listA.Add("");
                }
            }

            Worksheet activeWorksheet = app.ActiveSheet;

            for (int idxRow = 1; idxRow <= listA.Count; idxRow++)
            {
                ((Range)activeWorksheet.Cells[idxRow, 1]).Value2 = listA[idxRow - 1];
                ((Range)activeWorksheet.Cells[idxRow, 2]).Value2 = listB[idxRow - 1];
            }

            activeWorksheet.Columns.AutoFit();
        }
    }

    internal class MagentoDictionary
    {
        public int Line;
        public string ColumnA;
        public string ColumnB;
    }
}
