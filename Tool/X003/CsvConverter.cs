using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using ClosedXML.Excel;

namespace ExcelDiffTool
{
    public static class CsvConverter
    {
        // ワークブックをCSVとして保存

        public static List<string> SaveAsCsv(IXLWorkbook workbook, string path)
        {
            var list = new List<string>();
            var taskList = new List<Task>();

            foreach(var worksheet in workbook.Worksheets)
            {
                //worksheet.RangeUsed().DataType = XLDataType.Text;

                string name = Path.Combine(path, worksheet.Name);
                name += ".csv";

                list.Add(worksheet.Name);

                taskList.Add(
                    Task.Factory.StartNew(
                        () => { SaveAsCsv(worksheet, name); }
                        ));
                //SaveAsCsv(worksheet, name);
            }

            Task.WhenAll(taskList).Wait();

            return list;
        }


        // ワークシートをCSVとして保存
        public static void SaveAsCsv(IXLWorksheet worksheet, string filename)
        {
            var lines = new List<string>();
            var lastCellUsed = worksheet.LastCellUsed();

            if(lastCellUsed != null)
            {
                var rows = worksheet.Rows(1, lastCellUsed.Address.RowNumber);
                foreach (var row in rows)
                {
                    var values = new List<string>();
                    var cells = row.Cells(1, lastCellUsed.Address.ColumnNumber);
                    
                    foreach (var cell in cells)
                    {
                        if( cell.HasFormula) // エラーになるので対応
                        {
                            values.Add(cell.ValueCached.ToString());
                        }
                        else
                        {
                            values.Add(cell.Value.ToString());
                        }
                    }

                    string line = string.Join(";", values);
                    lines.Add(line);
                }
            }

            File.WriteAllLines(filename, lines);
        }

        private static string GetCellText(IXLCell cell)
        {
            string text = " ";

            if(cell != null)
            {
                text = cell.Value.ToString();
            }

            return text;
        }

        // ファイルの行リスト取得
        public static List<string> ReadCsvLines(string filename)
        {
            if(filename.IndexOf(".csv") == -1)
            {
                filename += ".csv";
            }

            var list = new List<string>();

            list.AddRange(File.ReadLines(filename));

            return list;
        }


        // CSVの行をカラムに分割
        public static List<string> GetCsvValues(string csvLine)
        {
            var list = new List<string>();

            list.AddRange(csvLine.Split(';'));

            return list;
        }
    }
}
