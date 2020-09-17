using System;
using System.IO;
using CsvToExcel.Models;

namespace CsvToExcel
{
    public class Program
    {
        /// <summary>
        /// エントリポイント
        /// </summary>
        /// <param name="args">$1:csvDirPath, $2:excelDirPath, $3:excelFileNameBase</param>
        public static void Main(string[] args)
        {
            // HACK: 引数この形でいいのか考える
            // 引数の取得とチェック
#if DEBUG
            var csvDirPath = @"./Sample";
            var excepDirPath = @"./Sample";
            var excelBaseName = @"Excel";
#else
            if (args.Length != 3)
            {
                throw new Exception("引数が不正です（$1:csvDirPath, $2:excelFilePath, $3:excelFileNameBase）。");
            }
            var csvDirPath = args[0];
            var excepDirPath = args[1];
            var excelBaseName = args[2];
#endif

            if (!Directory.Exists(csvDirPath))
            {
                throw new Exception("ディレクトリが見つかりません。" + csvDirPath);
            }

            if (!Directory.Exists(excepDirPath))
            {
                throw new Exception("ディレクトリが見つかりません。" + excepDirPath);
            }

            var timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
            var extension = ".xlsx";
            var excelFilePath = Path.Combine(excepDirPath, excelBaseName + timestamp + extension);

            // 定義の読み込み
            var csvDef = CsvDef.ReadCsvDef();
            var excelDef = ExcelDef.ReadExcelDef();

            // csvの読み込み
            var csvList = Csv.ReadAllCsv(csvDirPath, csvDef);

            // excel出力
            var creator = new Creator(excelFilePath, excelDef, csvList);
            creator.Create();
        }
    }
}
