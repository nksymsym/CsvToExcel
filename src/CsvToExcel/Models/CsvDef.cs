using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Xml.Serialization;

namespace CsvToExcel.Models
{
    /// <summary>
    /// csvファイルの構造を定義
    /// </summary>
    [XmlRoot("CsvDef")]
    public class CsvDef
    {
        /// <summary>
        /// ヘッダー行があるかどうか
        /// </summary>
        [XmlElement("HasHeader")]
        public bool HasHeader { get; set; }

        /// <summary>
        /// フッター行があるかどうか
        /// </summary>
        [XmlElement("HasFooter")]
        public bool HasFooter { get; set; }

        /// <summary>
        /// 読み飛ばす行番号
        /// </summary>
        /// <remarks>1はじまり</remarks>
        [XmlElement("SkipRowNumber")]
        public List<int> SkipRowNumbers { get; set; }

        /// <summary>
        /// 文字エンコーディング（文字列）
        /// </summary>
        [XmlElement("Encoding")]
        public string Encoding { get; set; }

        /// <summary>
        /// 区切り文字
        /// </summary>
        public string[] Separator { get; set; } = { "\t" };

        /// <summary>
        /// 設定ファイルの読み込み
        /// </summary>
        /// <returns>読み込み結果</returns>
        public static CsvDef ReadCsvDef()
        {
            var execDir = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
            var defPath = Path.Combine(execDir, @"./Config/CsvDef.xml");

            using (var file = new FileStream(defPath, FileMode.Open))
            {
                var serializer = new XmlSerializer(typeof(CsvDef));
                var def = (CsvDef)serializer.Deserialize(file);
                Check(def);
                return def;
            }
        }

        /// <summary>
        /// 読み込み結果のチェック
        /// </summary>
        /// <param name="def">読み込み結果</param>
        /// <remarks>エラー時は例外を投げる</remarks>
        private static void Check(CsvDef def)
        {
            try
            {
                var enc = System.Text.Encoding.GetEncoding(def.Encoding);
            }
            catch (ArgumentException e)
            {
                throw new Exception("Encodingの設定が不正です。", e);
            }
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <remarks>ReadCsvDefで初期化する</remarks>
        private CsvDef()
        {
        }
    }
}
