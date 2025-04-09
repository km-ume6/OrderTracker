using ClosedXML.Excel;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Data.SqlClient;

namespace OrderTracker
{
    internal class Program
    {
        class OrderElement
        {
            public string Name { get; set; } = string.Empty;
            public string Cell { get; set; } = string.Empty;
            public string Value { get; set; } = string.Empty;
            public int Type { get; set; } // 0:文字列, 1:日付, 2:日時, 3:数値
        }

        static void ReadOrderelement(List<OrderElement> elements, string xlPath)
        {
            // Excelファイルを開く
            using (var workbook = new XLWorkbook(xlPath))
            {
                var worksheet = workbook.Worksheet("受注票");
                {
                    foreach (var element in elements)
                    {
                        if (element.Cell != "")
                        {
                            // セルの値を取得
                            element.Value = worksheet.Cell(element.Cell).Value.ToString(); // 明示的な変換を追加
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 受注票のデータをデータベースに格納する
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("必須引数が指定されていません。");
                return;
            }

            if (!Directory.Exists(args[0]))
            {
                Console.WriteLine($"指定されたフォルダが存在しません: {args[0]}");
                return;
            }

            // ２番目のコマンドライン引数は日時  
            if (!DateTime.TryParse(args[1], out DateTime dateTime))
            {
                Console.WriteLine("無効な日時が指定されました。");
                return;
            }

            string thirdArg = args.Length >= 3 ? args[2] : "";
            if (thirdArg != "")
            {
                // 標準出力をリダイレクトする
                using (StreamWriter writer = new StreamWriter(thirdArg))
                {
                    Console.SetOut(writer);
                    Console.WriteLine($"出力先: {thirdArg}");
                    MainFunction(args[0], dateTime);
                }
            }
            else
            {
                // 標準出力をリダイレクトしない
                MainFunction(args[0], dateTime);
            }
        }

        static void MainFunction(string firstArg, DateTime dateTime)
        {
            Console.WriteLine($"指定されたフォルダ: {firstArg}");
            Console.WriteLine($"指定された日時: {dateTime.Date}");

            // キー入力を監視するためのタスクを開始
            CancellationTokenSource cts = new CancellationTokenSource();
            Task.Run(() =>
            {
                Console.WriteLine("キーを押すと処理を中断します...");
                Console.ReadKey();
                cts.Cancel();
            });

            List<OrderElement> DataPattern1 = new()
            {
                new OrderElement { Name = "受注番号", Cell = "AB3", Value = "", Type = 0 },
                new OrderElement { Name = "受注日", Cell = "C10", Value = "", Type = 1 },
                new OrderElement { Name = "取引先", Cell = "G10", Value = "", Type = 0 },
                new OrderElement { Name = "出荷日", Cell = "AE10", Value = "", Type = 1 },
                new OrderElement { Name = "注文番号", Cell = "G14", Value = "", Type = 0 },
                new OrderElement { Name = "商品コード", Cell = "C16", Value = "", Type = 0 },
                new OrderElement { Name = "製品番号", Cell = "O16", Value = "", Type = 0 },
                new OrderElement { Name = "製造指示書番号", Cell = "AA16", Value = "", Type = 0 },
                new OrderElement { Name = "品名仕様", Cell = "C19", Value = "", Type = 0 },
                new OrderElement { Name = "ラベル表記", Cell = "C23", Value = "", Type = 0 },
                new OrderElement { Name = "数量", Cell = "S23", Value = "0", Type = 0 },
                new OrderElement { Name = "ファイル名", Cell = "", Value = "", Type = 0 },
                new OrderElement { Name = "ファイル更新日時", Cell = "", Value = "", Type = 2 }
            };

            string connectionString = "Server=192.168.11.15;Database=Order;User Id=SangoKENSA;Password=227663m2;TrustServerCertificate=True;";
            string selectQuery = string.Empty, updateQuery = string.Empty, insertQuery = string.Empty;
            OrderElement? targetElement = null;

            try
            {
                DateTime timeStamp = DateTime.Now;
                var directories = Directory.EnumerateDirectories(firstArg);
                foreach (string directory in directories)
                {
                    if (cts.Token.IsCancellationRequested)
                    {
                        break;
                    }

                    Console.WriteLine(directory);

                    timeStamp = Directory.GetLastWriteTime(directory);
                    if (timeStamp >= dateTime)
                    {
                        try
                        {
                            // データベース接続
                            using (SqlConnection connection = new SqlConnection(connectionString))
                            {
                                connection.Open();

                                var files = Directory.EnumerateFiles(directory, "*.xls?", System.IO.SearchOption.AllDirectories);
                                foreach (string file in files)
                                {
                                    if (cts.Token.IsCancellationRequested)
                                    {
                                        break;
                                    }

                                    DateTime fileTimeStamp = File.GetLastWriteTime(file);
                                    if (fileTimeStamp >= dateTime)
                                    {
                                        Console.Write(file + " -> ");

                                        ReadOrderelement(DataPattern1, file);

                                        targetElement = DataPattern1.FirstOrDefault(item => item.Name == "ファイル名");
                                        if (targetElement != null)
                                        {
                                            targetElement.Value = file;
                                        }

                                        targetElement = DataPattern1.FirstOrDefault(item => item.Name == "ファイル更新日時");
                                        if (targetElement != null)
                                        {
                                            targetElement.Value = fileTimeStamp.ToString("yyyy/MM/dd HH:mm:ss");
                                        }

                                        selectQuery = $"SELECT * FROM 受注票 WHERE ファイル名 = '{file}'";

                                        // SQLコマンドの実行
                                        using (SqlCommand command = new SqlCommand(selectQuery, connection))
                                        {
                                            using (SqlDataReader reader = command.ExecuteReader())
                                            {
                                                // 結果を読み取る
                                                if (reader.Read())
                                                {
                                                    DateTime dbTimeStamp = reader.GetDateTime(reader.GetOrdinal("ファイル更新日時"));

                                                    reader.Close();

                                                    // ミリ秒以下を切り捨てる
                                                    DateTime fileTimeStampRounded = new DateTime(fileTimeStamp.Ticks - (fileTimeStamp.Ticks % TimeSpan.TicksPerSecond));
                                                    DateTime dbTimeStampRounded = new DateTime(dbTimeStamp.Ticks - (dbTimeStamp.Ticks % TimeSpan.TicksPerSecond));

                                                    // 秒単位で比較
                                                    if (fileTimeStampRounded > dbTimeStampRounded)
                                                    {
                                                        // データレコードを更新する
                                                        int i = 0;
                                                        updateQuery = $"UPDATE 受注票 SET ";
                                                        foreach (var item in DataPattern1)
                                                        {
                                                            updateQuery += $"{item.Name} = @dp1_{i++}, ";
                                                        }
                                                        updateQuery = updateQuery.TrimEnd(',', ' ') + $" WHERE ファイル名 = @FileName";

                                                        ExecuteSQL(connection, updateQuery, DataPattern1, "@FileName", file);

                                                        Console.WriteLine("データベース更新しました。");
                                                    }
                                                    else
                                                    {
                                                        Console.WriteLine("データベース変更なし");
                                                    }
                                                }
                                                else
                                                {
                                                    reader.Close();

                                                    int i = 0;
                                                    insertQuery = $"INSERT INTO 受注票 (";
                                                    foreach (var item in DataPattern1)
                                                    {
                                                        insertQuery += $"{item.Name}, ";
                                                    }
                                                    insertQuery = insertQuery.TrimEnd(',', ' ') + ") VALUES (";
                                                    foreach (var item in DataPattern1)
                                                    {
                                                        insertQuery += $"@dp1_{i++}, ";
                                                    }
                                                    insertQuery = insertQuery.TrimEnd(',', ' ') + ")";

                                                    ExecuteSQL(connection, insertQuery, DataPattern1);

                                                    Console.WriteLine("データベース追加しました。");
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        catch (SqlException ex)
                        {
                            Console.WriteLine($"SQLエラー: {ex.Message}");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"エラー: {ex.Message}");
                        }
                    }
                }

                static void ExecuteSQL(SqlConnection connection, string query, List<OrderElement> DataPattern, string op1 = "", string op2 = "")
                {
                    using (SqlCommand sqlCommand = new SqlCommand(query, connection))
                    {
                        DateTime dateTime;
                        int index = 0;
                        foreach (var item in DataPattern)
                        {
                            switch (item.Type)
                            {
                                case 1:
                                case 2:
                                    if (DateTime.TryParse(item.Value, out dateTime))
                                    {
                                        sqlCommand.Parameters.AddWithValue($"@dp1_{index++}", item.Type == 1 ? dateTime.Date : dateTime);
                                    }
                                    else
                                    {
                                        sqlCommand.Parameters.AddWithValue($"@dp1_{index++}", DateTime.Parse(item.Type == 1 ? "1969/8/21" : "1969/8/21 8:30"));
                                    }
                                    break;
                                case 3:
                                    sqlCommand.Parameters.AddWithValue($"@dp1_{index++}", int.Parse(item.Value));
                                    break;
                                default:
                                    sqlCommand.Parameters.AddWithValue($"@dp1_{index++}", item.Value);
                                    break;
                            }
                        }

                        // オプションパラメータを追加
                        if (op1 != "" && op2 != "")
                        {
                            sqlCommand.Parameters.AddWithValue(op1, op2);
                        }

                        sqlCommand.ExecuteNonQuery();
                    }
                }
            }
            catch (OperationCanceledException)
            {
                Console.WriteLine("処理がキャンセルされました。");
            }
        }
    }
}
