using System;

namespace SimpleExcelXml
{
    /// <summary>
    /// プロパティに表示するセル位置を文字列(A1等)で与える
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class CellAttribute : Attribute
    {
        /// <summary>
        /// セルの座標群
        /// </summary>
        public string[] Positions { get; set; }

        /// <summary>
        /// 出力フォーマット
        /// </summary>
        public string Format { get; set; }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="format">出力フォーマット</param>
        /// <param name="positions">座標群</param>
        public CellAttribute(string format, params string[] positions)
        {
            this.Positions = positions;
            this.Format = format;
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="position">座標</param>
        public CellAttribute(string position)
        {
            this.Positions = new string[] { position };
        }
    }
}
