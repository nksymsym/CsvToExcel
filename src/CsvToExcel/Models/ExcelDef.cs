using System.IO;
using System.Reflection;
using System.Xml.Serialization;

namespace CsvToExcel.Models
{
    /// <summary>
    /// excelファイルの設定を定義
    /// </summary>
    [XmlRoot("ExcelDef")]
    public class ExcelDef
    {
        /// <summary>
        /// シートを分けるかどうか
        /// </summary>
        [XmlElement("IsMultipleSheets")]
        public bool IsMultipleSheets { get; set; }

        /// <summary>
        /// タイトル行があるかどうか
        /// </summary>
        [XmlElement("HasTitle")]
        public bool HasTitle { get; set; }

        /// <summary>
        /// タイトルに拡張子を含めるかどうか
        /// </summary>
        [XmlElement("HasTitleExt")]
        public bool HasTitleExt { get; set; }

        /// <summary>
        /// タイトル行を太字にするかどうか
        /// </summary>
        [XmlElement("IsTitleBold")]
        public bool IsTitleBold { get; set; }

        /// <summary>
        /// ヘッダー行を太字にするかどうか
        /// </summary>
        [XmlElement("IsHeaderBold")]
        public bool IsHeaderBold { get; set; }

        /// <summary>
        /// データ行を太字にするかどうか
        /// </summary>
        [XmlElement("IsDataBold")]
        public bool IsDataBold { get; set; }

        /// <summary>
        /// フッター行を太字にするかどうか
        /// </summary>
        [XmlElement("IsFooterBold")]
        public bool IsFooterBold { get; set; }

        /// <summary>
        /// ヘッダー行の背景色
        /// </summary>
        [XmlElement("HeaderBgColor")]
        public string HeaderBgColor { get; set; }

        /// <summary>
        /// データ行の背景色
        /// </summary>
        [XmlElement("DataBgColor")]
        public string DataBgColor { get; set; }

        /// <summary>
        /// フッター行の背景色
        /// </summary>
        [XmlElement("FooterBgColor")]
        public string FooterBgColor { get; set; }

        /// <summary>
        /// 前に何列空白を入れるか
        /// </summary>
        [XmlElement("LeadingColumns")]
        public int LeadingColumns { get; set; }

        /// <summary>
        /// 前に何行空白を入れるか
        /// </summary>
        [XmlElement("LeadingRows")]
        public int LeadingRows { get; set; }

        /// <summary>
        /// 設定ファイルの読み込み
        /// </summary>
        /// <returns>読み込み結果</returns>
        public static ExcelDef ReadExcelDef()
        {
            var execDir = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
            var defPath = Path.Combine(execDir, @"./Config/ExcelDef.xml");

            using (var file = new FileStream(defPath, FileMode.Open))
            {
                var serializer = new XmlSerializer(typeof(ExcelDef));
                var def = (ExcelDef)serializer.Deserialize(file);
                return def;
            }
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <remarks>ReadExcelDefで初期化する</remarks>
        private ExcelDef()
        {
        }
    }
}
