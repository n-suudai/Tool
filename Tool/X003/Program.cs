using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace ExcelDiffTool
{
    class Program
    {

        static void Main(string[] args)
        {
            Log.Init("ExcelDiffTool_New_Log.txt");
#if !DEBUG
            try
#endif
            {
                string[] files = { "", "" };
                string[] names = { "src", "dst" };
                List<string>[] sheetLists = { null, null };

                // ファイル入力待ち
                {
                    if (args.Length == 2)
                    {
                        files[0] = args[0];
                        files[1] = args[1];
                    }
                    else
                    {
                        Log.WriteLine("ファイル１をドラッグアンドドロップしてください");
                        files[0] = Log.ReadLine();

                        Log.WriteLine("ファイル2をドラッグアンドドロップしてください");
                        files[1] = Log.ReadLine();
                    }
                }

                // CSVにコンバート
                for (int i = 0; i < names.Length; ++i)
                {
                    Directory.CreateDirectory(names[i]);

                    Log.WriteLine(string.Format("{0} をCSVに変換しています。", files[i]));
                    sheetLists[i] = CsvConverter.SaveAsCsv(new XLWorkbook(files[i]), names[i]);
                }

                // 出力用ワークシート
                var outWb = new XLWorkbook();

                // シート差分 ワークシートの追加、削除を検出

                bool sheetDiffFound = false;

                sheetLists[0].Sort();
                sheetLists[1].Sort();

                var sheetDiffList = Diff.Execute(sheetLists[0], sheetLists[1]);
                foreach (var sheetDiff in sheetDiffList)
                {
                    var sheet = outWb.AddWorksheet(sheetDiff.Text);

                    // シートのステータスによって色を変更
                    if (sheetDiff.Status == Diff.Status.Add)
                    {
                        Log.WriteLine(string.Format("「{0}」シートが追加されました。", sheet.Name));

                        sheetDiffFound = true;
                        sheet.TabColor = XLColor.LightGreen;

                        var lines = CsvConverter.ReadCsvLines(Path.Combine(names[1], sheet.Name));

                        for (int rowIndex = 0; rowIndex < lines.Count; ++rowIndex)
                        {
                            IXLRow row = sheet.Row(rowIndex + 1);

                            // 行番号出力
                            row.Cell(1).Value = rowIndex + 1;

                            var columnList = CsvConverter.GetCsvValues(lines[rowIndex]);

                            for (int columnIndex = 0; columnIndex < columnList.Count; ++columnIndex)
                            {// カラム出力
                                row.Cell(columnIndex + 2).Value = columnList[columnIndex];
                            }

                            var cells = row.Cells(1, columnList.Count + 1);
                            cells.Style.Fill.PatternType = XLFillPatternValues.Solid;
                            cells.Style.Fill.BackgroundColor = XLColor.LightGreen;
                        }
                    }
                    else if (sheetDiff.Status == Diff.Status.Delete)
                    {
                        Log.WriteLine(string.Format("「{0}」シートが削除されました。", sheet.Name));

                        sheetDiffFound = true;
                        sheet.TabColor = XLColor.LightCoral;

                        var lines = CsvConverter.ReadCsvLines(Path.Combine(names[0], sheet.Name));

                        for (int rowIndex = 0; rowIndex < lines.Count; ++rowIndex)
                        {
                            IXLRow row = sheet.Row(rowIndex + 1);

                            // 行番号出力
                            row.Cell(1).Value = rowIndex + 1;

                            var columnList = CsvConverter.GetCsvValues(lines[rowIndex]);

                            for (int columnIndex = 0; columnIndex < columnList.Count; ++columnIndex)
                            {// カラム出力
                                row.Cell(columnIndex + 2).Value = columnList[columnIndex];
                            }

                            var cells = row.Cells(1, columnList.Count + 1);
                            cells.Style.Fill.PatternType = XLFillPatternValues.Solid;
                            cells.Style.Fill.BackgroundColor = XLColor.LightCoral;
                        }
                    }
                    else
                    { // 両方に存在するシートなので、差分を検出する

                        Log.WriteLine(string.Format("「{0}」シートの差分を検出中です。", sheet.Name));

                        // 行差分 行の追加、削除を検出
                        List<Diff.Line> rowDiffList = null;
                        {
                            var src = CsvConverter.ReadCsvLines(Path.Combine(names[0], sheet.Name));
                            var dst = CsvConverter.ReadCsvLines(Path.Combine(names[1], sheet.Name));
                            rowDiffList = Diff.Execute(src, dst);
                        }

                        bool rowDiffFound = false;
                        int addCount = 0;
                        int deleteCount = 0;
                        for (int rowIndex = 0; rowIndex < rowDiffList.Count; ++rowIndex)
                        {
                            // とりあえず全部出力
                            var rowDiff = rowDiffList[rowIndex];
                            IXLRow row = sheet.Row(rowIndex + 1);

                            // 行番号出力
                            {
                                var cell = row.Cell(1);
                                cell.Value = rowDiff.Index + 1;
                                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                            }

                            var columnList = CsvConverter.GetCsvValues(rowDiff.Text);
                            for (int columnIndex = 0; columnIndex < columnList.Count; ++columnIndex)
                            { // カラム出力
                                row.Cell(columnIndex + 2).Value = columnList[columnIndex];
                            }

                            if (rowDiff.Status == Diff.Status.Add)
                            {
                                addCount++;
                                rowDiffFound = true;
                                var cells = row.Cells(1, columnList.Count + 1);
                                cells.Style.Fill.PatternType = XLFillPatternValues.Solid;
                                cells.Style.Fill.BackgroundColor = XLColor.LightGreen;
                            }
                            else if (rowDiff.Status == Diff.Status.Delete)
                            {
                                deleteCount++;
                                rowDiffFound = true;
                                var cells = row.Cells(1, columnList.Count + 1);
                                cells.Style.Fill.PatternType = XLFillPatternValues.Solid;
                                cells.Style.Fill.BackgroundColor = XLColor.LightCoral;
                            }
                        }

                        if (rowDiffFound)
                        {
                            sheetDiffFound = true;
                            Log.WriteLine("差分が検出されました。");
                            Log.WriteLine(string.Format("追加された行数 : {0}", addCount));
                            Log.WriteLine(string.Format("削除された行数 : {0}", deleteCount));
                        }
                        else
                        {
                            sheet.Delete(); // 差分は無いので出力に含める必要が無い
                        }
                    }
                }

                if (!sheetDiffFound)
                {
                    Log.WriteLine("シート差分はありません。");
                }
                else
                {
                    Log.WriteLine("保存中です。");
                    outWb.SaveAs("ExcelDiffTool_結果.xlsx");
                }
            }
#if !DEBUG
            catch (Exception e)
            {
                Log.WriteLine("=== 例外が発生!! ===");
                Log.WriteLine(e.ToString());
            }
#endif

            Log.WriteLine("終了しました。");
            Log.Term();
            Console.ReadLine();
        }
    }
}
