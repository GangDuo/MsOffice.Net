using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace MsOffice
{
    public class MsExcel
    {
        public static string SaveAs(string filename, Excel.XlFileFormat xlType = Excel.XlFileFormat.xlOpenXMLWorkbook)
        {
            var path = string.Empty;
            Transaction(filename, (wkbk) =>
            {
                wkbk.SaveAs(GetPathWithoutExtension(filename), (int)xlType);
                path = wkbk.FullName;
            });
            return path;
        }

        public static string SaveAsCsv(string filename, bool isFirstLineEliminated = true)
        {
            var path = string.Empty;
            Transaction(filename, (wkbk) =>
            {
                if (isFirstLineEliminated)
                {
                    Excel.Worksheet sheet = wkbk.Sheets[1];
                    sheet.get_Range("1:1").Delete();
                }
                wkbk.SaveAs(GetPathWithoutExtension(filename), Excel.XlFileFormat.xlCSV, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                path = wkbk.FullName;
            });
            return path;
        }

        public static string SaveAsPdf(string filename)
        {
            var path = string.Empty;
            Transaction(filename, (wkbk) =>
            {
                Excel.Worksheet sheet = wkbk.Sheets[1];
                path = Path.Combine(Path.GetDirectoryName(filename), Path.GetFileNameWithoutExtension(filename) + ".pdf");
                sheet.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, path, Excel.XlFixedFormatQuality.xlQualityStandard);
            });
            return path;
        }

        // シート名、行、列、値
        public static void SetValue(string filename, List<Tuple<object, int, int, object>> xs)
        {
            Transaction(filename, (wkbk) =>
            {
                foreach (var x in xs)
                {
                    Debug.Assert(ContainsWorksheetName(wkbk.Sheets, x.Item1.ToString()));
                    Excel.Worksheet sheet = wkbk.Sheets[x.Item1];
                    sheet.Cells[x.Item2, x.Item3] = x.Item4;                    
                }
                wkbk.Save();
            });
        }

        // シート名またはシート番号（1始まり）、削除する行数
        public static void DeleteTail(string filename, List<Tuple<object, int>> xs)
        {
            Transaction(filename, (wkbk) =>
            {
                foreach (var x in xs)
                {
                    for (var i = 0; i < x.Item2; ++i)
                    {
                        Excel.Worksheet s = wkbk.Sheets[x.Item1];
                        s.Rows[s.Cells[1, 1].End[Excel.XlDirection.xlDown].Row].Delete();
                    }
                }
                wkbk.Save();
            });
        }

        public static void DeleteRows(string filename, List<Tuple<object, string>> xs)
        {
            Transaction(filename, (wkbk) =>
            {
                foreach (var x in xs)
                {
                    wkbk.Sheets[x.Item1].Rows[x.Item2].Delete();
                }
                wkbk.Save();
            });
        }

        // シート名、列範囲
        public static void DeleteColumns(string filename, List<Tuple<object, string>> xs)
        {
            Transaction(filename, (wkbk) =>
            {
                foreach (var x in xs)
                {
                    wkbk.Sheets[x.Item1].Columns[x.Item2].Delete();
                }
                wkbk.Save();
            });
        }

        private static bool ContainsWorksheetName(Excel.Sheets sheets, string name)
        {
            var xs = new List<string>();
            foreach (Excel.Worksheet s in sheets)
            {
                xs.Add(s.Name);
            }
            return xs.Contains(name);
        }

        private static string GetPathWithoutExtension(string path)
        {
            return Path.Combine(Path.GetDirectoryName(path), Path.GetFileNameWithoutExtension(path));
        }

        private static void Transaction(string filename, Action<Excel.Workbook> callback)
        {
            Excel.Application excelApp = null;
            Excel.Workbook wkbk = null;

            try
            {
                excelApp = new Excel.Application();
#if DEBUG
                excelApp.Visible = true;
#else
                excelApp.Visible = false;
#endif
                excelApp.DisplayAlerts = false;
                wkbk = excelApp.Workbooks.Open(filename);
                callback(wkbk);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                if (null != wkbk)
                {
                    wkbk.Close();
                    wkbk = null;
                }

                // Close Excel.
                if (null != excelApp)
                {
                    excelApp.Quit();
                    excelApp = null;
                }
            }
        }
    }
}
