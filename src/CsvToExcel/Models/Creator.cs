using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;

namespace CsvToExcel.Models
{
    /// <summary>
    /// Excel作成
    /// </summary>
    public class Creator
    {
        /// <summary>
        /// 作成ファイルパス
        /// </summary>
        private readonly string filePath;

        /// <summary>
        /// excel定義
        /// </summary>
        private readonly ExcelDef def;

        /// <summary>
        /// csvデータ
        /// </summary>
        private readonly IReadOnlyCollection<Csv> csvList;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="excelFilePath">作成ファイルパス</param>
        /// <param name="excelDef">excel定義</param>
        /// <param name="csvList">csvデータ</param>
        public Creator(string excelFilePath, ExcelDef excelDef, IReadOnlyCollection<Csv> csvList)
        {
            filePath = excelFilePath;
            def = excelDef;
            this.csvList = csvList;
        }

        /// <summary>
        /// Excelファイル作成
        /// </summary>
        public void Create()
        {
            // データがない場合は作成しない
            if (csvList.Count == 0)
            {
                return;
            }

            // ブックを作成
            using (var workbook = new XLWorkbook())
            {
                if (def.IsMultipleSheets)
                {
                    // 1シートずつ作成
                    CreateSheetForOne(workbook);
                }
                else
                {
                    // まとめて1シートに作成
                    CreateSheetForAll(workbook);
                }

                // ファイルに保存
                workbook.SaveAs(filePath);
            }
        }

        private void CreateSheetForAll(XLWorkbook workbook)
        {
            // HACK: ほとんど同じなので共通化したい気もする
            var worksheet = workbook.Worksheets.Add("Data");

            var row = def.LeadingRows;
            var col = def.LeadingColumns;

            foreach (var csv in csvList)
            {
                var title = csv.FileName;
                if (!def.HasTitleExt)
                {
                    title = Path.GetFileNameWithoutExtension(csv.FileName);
                }

                // タイトル
                row++;
                PasteLine(worksheet, row, col, new[] { title }, def.IsTitleBold, "NoColor", false);

                // ヘッダー
                row++;
                PasteLine(worksheet, row, col, csv.Header, def.IsHeaderBold, def.HeaderBgColor, true);

                foreach (var data in csv.DataList)
                {
                    // データ
                    row++;
                    PasteLine(worksheet, row, col, data, def.IsDataBold, def.DataBgColor, true);
                }

                // フッター
                row++;
                PasteLine(worksheet, row, col, csv.Footer, def.IsFooterBold, def.FooterBgColor, true);

                // 空行を入れる
                row++;
            }
        }

        private void CreateSheetForOne(XLWorkbook workbook)
        {
            int sheetNo = 0;
            foreach (var csv in csvList)
            {
                sheetNo++;

                var title = csv.FileName;
                if (!def.HasTitleExt)
                {
                    title = Path.GetFileNameWithoutExtension(csv.FileName);
                }

                // TODO: 31文字対応
                // TODO: 重複対応（先頭31文字が同じ場合）
                string sheetName = title;
                var worksheet = workbook.Worksheets.Add(sheetName);

                var row = def.LeadingRows;
                var col = def.LeadingColumns;

                // タイトル
                row++;
                PasteLine(worksheet, row, col, new[] { title }, def.IsTitleBold, "NoColor", false);

                // ヘッダー
                row++;
                PasteLine(worksheet, row, col, csv.Header, def.IsHeaderBold, def.HeaderBgColor, true);

                foreach (var data in csv.DataList)
                {
                    // データ
                    row++;
                    PasteLine(worksheet, row, col, data, def.IsDataBold, def.DataBgColor, true);
                }

                // フッター
                row++;
                PasteLine(worksheet, row, col, csv.Footer, def.IsFooterBold, def.FooterBgColor, true);
            }
        }

        private void PasteLine(
            IXLWorksheet worksheet, int row, int col,
            IReadOnlyCollection<string> line, bool isBold, string bgColor, bool hasBoeder)
        {
            foreach (var item in line)
            {
                col++;

                var cell = worksheet.Cell(row, col);
                cell.Value = item;

                var style = cell.Style;
                // HACK:事前チェック
                style.Fill.BackgroundColor = XLColor.FromName(bgColor);
                style.Font.Bold = isBold;
                if (hasBoeder)
                {
                    style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                }
            }
        }
    }
}
