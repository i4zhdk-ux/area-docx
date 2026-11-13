using System;
using System.Diagnostics;
using System.IO;
using NPOI.XWPF.UserModel;
using System.Globalization;

namespace WordTableReader
{
    class Program
    {
        static void Main(string[] args)
        {
            // 1. 引数のチェック
            if (args.Length == 0)
            {
                Console.WriteLine("使用法: wordApp.exe <ファイルパス>");
                return;
            }

            string filePath = args[0];

            if (!File.Exists(filePath))
            {
                Console.WriteLine($"エラー: ファイルが見つかりません: {filePath}");
                return;
            }

            // 2. ストップウォッチの開始
            Stopwatch sw = new Stopwatch();
            sw.Start();

            DateOnly today = DateOnly.FromDateTime(DateTime.Now);
            DateOnly nextYear = today.AddYears(1);
            // 4か月前の日付を取得
            DateOnly fourMonthsAgo = today.AddMonths(-4);
            var list = new List<(string number, DateOnly date)>();
            var fourMonthsAgoList = new List<string>();
            var noExistNumberList = new List<string>();

            try
            {
                Console.WriteLine($"{Path.GetFileName(filePath)} を読み取り中...");

                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    // Wordドキュメントの読み込み
                    XWPFDocument doc = new XWPFDocument(fs);

                    // 3. 表（テーブル）の読み取り
                    int tableCount = 1;
                    foreach (var table in doc.Tables)
                    {
                        if(table.Rows.Count < 2)
                        {
                            continue;
                        }
                        Console.WriteLine($"{tableCount++}ページ目を読み取り中");
                        for (int i = 2; i < table.Rows.Count; i++)
                        {
                            var row = table.Rows[i];
                            if (i % 2 == 0) 
                            {
                                continue;
                            }
                            
                            var cells = row.GetTableCells();
                            for (int j = 2; j < cells.Count; j++)
                            {
                                string cellValue = cells[j].GetText().Trim();
                                var preRow = table.Rows[i - 1];
                                var preCell = preRow.GetTableCells();
                                if (preCell[0].GetText().Trim() == "")
                                {
                                    break;
                                }
                                if (j == 2 && cellValue == "") {
                                    if (!parseDate(preCell[1].GetText().Trim())) {
                                        Console.WriteLine($"エラーが発生しました。区域番号 {preCell[0].GetText().Trim()} の日付 {preCell[1].GetText().Trim()} が間違っています。");
                                    } else {
                                        var number = preCell[0].GetText().Trim();
                                        var date = DateOnly.Parse(preCell[1].GetText().Trim());
                                        list.Add((number, date));
                                    }
                                    break; 
                                } else if (j % 2 == 1 && cellValue == "") {
                                    var number = preCell[0].GetText().Trim();
                                    var date = DateOnly.Parse(cells[j - 1].GetText().Trim());
                                    if (date <= fourMonthsAgo)
                                    {
                                        fourMonthsAgoList.Add(number);
                                    }
                                    list.Add((preCell[0].GetText().Trim(), nextYear));
                                    break;
                                } else if (cellValue == "") {
                                    if (!parseDate(cells[j - 1].GetText().Trim())) {
                                        Console.WriteLine($"エラーが発生しました。区域番号 {preCell[0].GetText().Trim()} の日付 {cells[j - 1].GetText().Trim()} が間違っています。");
                                    } else {
                                        var number = preCell[0].GetText().Trim();
                                        var date = DateOnly.Parse(cells[j - 1].GetText().Trim());
                                        list.Add((number, date));
                                    }
                                    break;
                                }
                            }
                        }
                    }
                }
                sortAndWriteResultsToCsv(list, nextYear);
                Console.WriteLine("*******抽出結果にない区域番号は、以下の通りです。*****");
                for(var i = 1; i <= 306; i++)
                {
                    var existNumber = list.Any(item => item.number == i.ToString());
                    if (!existNumber)
                    {
                        noExistNumberList.Add(i.ToString());
                    }
                }
                for(var i = 1; i <= 31; i++)
                {
                    var existNumber = list.Any(item => item.number == "T" + i.ToString());
                    if (!existNumber)
                    {
                        noExistNumberList.Add("T" + i.ToString());
                    }
                }
                if (noExistNumberList.Count == 0)
                {
                    Console.WriteLine("なし。");
                }
                foreach (var number in noExistNumberList)
                {
                    Console.WriteLine(number);
                }
                Console.WriteLine("*******貸出から4か月以上経った区域番号は、以下の通りです。*****");
                if (fourMonthsAgoList.Count == 0)
                {
                    Console.WriteLine("なし。");
                }
                foreach (var number in fourMonthsAgoList)
                {
                    Console.WriteLine(number);
                }

                // 4. 計測終了
                sw.Stop();
                Console.WriteLine("----------------------------------");
                Console.WriteLine($"処理完了。");
                Console.WriteLine($"実行時間: {sw.ElapsedMilliseconds} ms ({sw.Elapsed.TotalSeconds:F2} 秒)");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"エラーが発生しました: {ex.Message}");
            }
        }

        private static bool parseDate(string input) {
            string format = "yy/M/d";
            if (DateOnly.TryParseExact(input, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateOnly result))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private static void sortAndWriteResultsToCsv(List<(string number, DateOnly date)> list, DateOnly nextYear) {
            var sortedList = list.OrderBy(x => x.date).ToList();

            List<(string number, string date)> updatedList = sortedList.Select(x => 
            {
                string newStatus = x.date == nextYear ? "未返却" : x.date.ToString("yy/MM/dd");
                return (x.number, newStatus);
            }).ToList();
            var csvLines = updatedList.Select(x => $"{x.number},{x.date}");
            File.WriteAllLines("抽出結果.txt", csvLines);
        }
    }
}