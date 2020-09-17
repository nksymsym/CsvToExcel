using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace CsvToExcel.Models
{
    /// <summary>
    /// 各csvファイルの内容
    /// </summary>
    public class Csv
    {
        /// <summary>
        /// ファイル名
        /// </summary>
        public string FileName { get; private set; }

        /// <summary>
        /// ヘッダー行
        /// </summary>
        public IReadOnlyCollection<string> Header { get; private set; }

        /// <summary>
        /// データ行（全行）
        /// </summary>
        public IReadOnlyCollection<IReadOnlyCollection<string>> DataList { get; private set; }

        /// <summary>
        /// フッター行
        /// </summary>
        public IReadOnlyCollection<string> Footer { get; private set; }

        /// <summary>
        /// csvをすべて読み込む
        /// </summary>
        /// <param name="dirPath">csvの配置ディレクトリバス</param>
        /// <param name="def">csv定義</param>
        /// <returns>読み込み結果</returns>
        public static IReadOnlyCollection<Csv> ReadAllCsv(string dirPath, CsvDef def)
        {
            var csvList = new List<Csv>();

            foreach (var filePath in Directory.GetFiles(dirPath, "*.csv"))
            {
                csvList.Add(Csv.ReadCsv(filePath, def));
            }

            return csvList;
        }

        /// <summary>
        /// csvファイルの読み込み
        /// </summary>
        /// <returns>読み込み結果</returns>
        private static Csv ReadCsv(string filePath, CsvDef def)
        {
            // 初期化
            var csv = new Csv();
            var dataList = new List<List<string>>();
            csv.DataList = dataList;

            // 存在チェック
            if (!File.Exists(filePath))
            {
                throw new Exception("ファイルが見つかりません。" + filePath);
            }

            // ファイル名を設定
            csv.FileName = Path.GetFileName(filePath);

            // csv読み込み
            var enc = Encoding.GetEncoding(def.Encoding);
            using (var reader = new StreamReader(filePath, enc))
            {
                var rowNumber = 0;
                while (!reader.EndOfStream)
                {
                    // 1行読み込み
                    var line = reader.ReadLine();
                    rowNumber++;

                    // skip確認
                    if (def.SkipRowNumbers.Contains(rowNumber))
                    {
                        continue;
                    }

                    // lineを分割
                    var items = line.Split(def.Separator, StringSplitOptions.None).ToList();

                    // 先頭行はHeader
                    if (def.HasHeader && csv.Header == null)
                    {
                        csv.Header = items;
                        continue;
                    }

                    // 最終行はFooter
                    if (def.HasFooter && reader.EndOfStream)
                    {
                        csv.Footer = items;
                        continue;
                    }

                    // 中間行はData
                    dataList.Add(items);
                }
            }

            return csv;
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <remarks>ReadCsvで初期化する</remarks>
        private Csv()
        {
        }
    }
}
