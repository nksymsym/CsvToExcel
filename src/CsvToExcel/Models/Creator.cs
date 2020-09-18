using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Presentation;

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
                workbook.Style.Font.FontName = def.FontName;
                workbook.Style.Font.FontSize = def.FontSize;
                workbook.Style.NumberFormat.Format = "@";

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

        /// <summary>
        /// 全csvデータから1シートを作成
        /// </summary>
        /// <param name="workbook">ブック</param>
        private void CreateSheetForAll(XLWorkbook workbook)
        {
            var worksheet = workbook.Worksheets.Add("Data");

            var row = def.LeadingRows;
            var col = def.LeadingColumns;

            foreach (var csv in csvList)
            {
                string title = getTitle(csv);

                // csvデータを出力
                row = PasteCsv(csv, title, worksheet, row, col);

                // 空行を入れる
                row++;
            }
        }

        /// <summary>
        /// 全csvデータからcsvごとに1シート作成
        /// </summary>
        /// <param name="workbook">ブック</param>
        private void CreateSheetForOne(XLWorkbook workbook)
        {
            int sheetNo = 0;
            foreach (var csv in csvList)
            {
                sheetNo++;

                string title = getTitle(csv);

                // TODO: 31文字対応
                // TODO: 重複対応（先頭31文字が同じ場合）
                string sheetName = title;
                var worksheet = workbook.Worksheets.Add(sheetName);

                var row = def.LeadingRows;
                var col = def.LeadingColumns;

                // csvデータを出力
                row = PasteCsv(csv, title, worksheet, row, col);
            }
        }

        /// <summary>
        /// csvデータをセルに出力
        /// </summary>
        /// <param name="csv">csvデータ</param>
        /// <param name="title">タイトル</param>
        /// <param name="worksheet">シート</param>
        /// <param name="row">行番号</param>
        /// <param name="col">列番号</param>
        /// <returns></returns>
        private int PasteCsv(Csv csv, string title, IXLWorksheet worksheet, int row, int col)
        {

            // タイトル
            if (def.HasTitle)
            {
                row++;
                PasteLine(worksheet, row, col, new[] { title }, def.IsTitleBold, "NoColor", false);
            }

            // ヘッダー
            if (csv.Header != null && csv.Header.Count != 0)
            {
                row++;
                PasteLine(worksheet, row, col, csv.Header, def.IsHeaderBold, def.HeaderBgColor, true);
            }

            foreach (var data in csv.DataList)
            {
                // データ
                if (data != null && data.Count != 0)
                {
                    row++;
                    PasteLine(worksheet, row, col, data, def.IsDataBold, def.DataBgColor, true);
                }
            }

            // フッター
            if (csv.Footer != null && csv.Footer.Count != 0)
            {
                row++;
                PasteLine(worksheet, row, col, csv.Footer, def.IsFooterBold, def.FooterBgColor, true);
            }

            return row;
        }

        /// <summary>
        /// 1行分のデータをセルに出力
        /// </summary>
        /// <param name="worksheet">シート</param>
        /// <param name="row">行番号</param>
        /// <param name="col">列番号</param>
        /// <param name="line">行データ</param>
        /// <param name="isBold">太字にするか</param>
        /// <param name="bgColor">背景色</param>
        /// <param name="hasBoeder">外枠の罫線をつけるか</param>
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

        /// <summary>
        /// タイトル取得
        /// </summary>
        /// <param name="csv">csvデータ</param>
        /// <returns>タイトル</returns>
        private string getTitle(Csv csv)
        {
            var title = csv.FileName;
            if (!def.HasTitleExt)
            {
                title = Path.GetFileNameWithoutExtension(csv.FileName);
            }

            return title;
        }
    }
}
