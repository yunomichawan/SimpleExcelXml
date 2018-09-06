namespace SimpleExcelXml
{
    public class ExcelSampleObject
    {
        [Cell("{0}!", "A1")]
        public string Title
        {
            get
            {
                return this.CompanyName + " × " + this.Name;
            }
        }

        [Cell("A2")]
        public string CompanyName { get; set; }

        [Cell("商品名：{0}", "B2")]
        public string Name { get; set; }

        [Cell("C2")]
        public int Price { get; set; }

        [Cell("D2")]
        public string Introduction { get; set; }

        /// <summary>
        /// 出力したいセルが複数あって、フォーマットが存在しない場合はフォーマットに空を設定
        /// </summary>
        [Cell("", "E2", "E4")]
        public string Remarks { get; set; }

        /// <summary>
        /// 出力対象外のプロパティ
        /// Cell属性を適用していないプロパティは出力されません。
        /// </summary>
        public string Excluded { get; set; }
    }
}
