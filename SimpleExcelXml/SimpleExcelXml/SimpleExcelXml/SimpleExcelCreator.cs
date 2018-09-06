using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.IO;
using System.Linq;

namespace SimpleExcelXml
{
    /// <summary>
    /// Excelブック操作クラス
    /// </summary>
    public class SimpleExcelCreator
    {
        #region property

        /// <summary>
        /// 出力パス
        /// </summary>
        private string OutputPath { get; set; }

        /// <summary>
        /// テンプレートパス
        /// </summary>
        private string TempPath { get; set; }

        /// <summary>
        /// ワークブック
        /// </summary>
        private WorkbookPart WorkbookPart { get { return this.SpreadsheetDoc.WorkbookPart; } }

        /// <summary>
        /// ワークシート操作クラス
        /// </summary>
        public WorksheetControl WorksheetControl { get; set; }

        /// <summary>
        /// Excelの皿
        /// </summary>
        private SpreadsheetDocument SpreadsheetDoc { get; set; }

        #endregion

        #region constructor

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="outputPath">出力パス</param>
        public SimpleExcelCreator(string outputPath, SpreadsheetDocumentType type = SpreadsheetDocumentType.Workbook)
        {
            this.OutputPath = outputPath;
            this.SpreadsheetDoc = SpreadsheetDocument.Create(this.OutputPath, type);
            this.Init(true);
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="tempPath">テンプレートファイルパス</param>
        /// <param name="outputPath">出力パス</param>
        public SimpleExcelCreator(string tempPath, string outputPath)
        {
            this.OutputPath = outputPath;
            this.TempPath = tempPath;
            File.Copy(this.TempPath, this.OutputPath, true);
            this.SpreadsheetDoc = SpreadsheetDocument.Open(this.TempPath, true);
            this.Init(false);
        }

        #endregion

        /// <summary>
        /// 初期化
        /// </summary>
        /// <param name="spreadsheetDocument"></param>
        /// <param name="isNew"></param>
        private void Init(bool isNew)
        {
            if (isNew)
            {
                this.SpreadsheetDoc.AddWorkbookPart();
                this.WorkbookPart.Workbook = new Workbook();
                this.WorkbookPart.Workbook.AppendChild(new Sheets());
                this.WorksheetControl = new WorksheetControl(this.WorkbookPart);
            }
            else
            {
                this.WorksheetControl = new WorksheetControl(this.WorkbookPart);
            }
        }

        #region sheet

        /// <summary>
        /// シート追加
        /// </summary>
        /// <param name="name"></param>
        public void AddSheet(string name)
        {
            // 同名のシートが存在するかチェック
            if (this.WorkbookPart.Workbook.Sheets.Any(s => ((Sheet)s).Name.Value.Equals(name)))
                return;

            WorksheetPart worksheetPart = this.WorkbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());
            Sheet sheet = new Sheet
            {
                Id = this.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = (UInt32)(this.WorkbookPart.Workbook.Sheets.Count() + 1), // 0はNG
                Name = name,
            };
            this.WorkbookPart.Workbook.Sheets.Append(sheet);
            this.WorksheetControl.AddWorksheetObject(new WorksheetObject(worksheetPart, sheet));
        }

        /// <summary>
        /// シートの削除
        /// </summary>
        /// <param name="name"></param>
        public void RemoveSheet(string name)
        {
            Sheet sheet = this.WorkbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault(s => s.Name.Value.Equals(name));
            if (sheet != null)
            {
                this.WorkbookPart.Workbook.Sheets.RemoveChild(sheet);
                var worksheetPart = this.WorkbookPart.GetPartById(sheet.Id);
                this.WorksheetControl.RemoveWorksheetObject(worksheetPart);
                this.WorkbookPart.DeletePart(worksheetPart);
            }
        }

        /// <summary>
        /// シート選択
        /// </summary>
        /// <param name="index"></param>
        public WorksheetObject SelectSheet(int index)
        {
            return this.WorksheetControl.SetCurrent(index);
        }

        /// <summary>
        /// シート選択
        /// </summary>
        /// <param name="sheetName"></param>
        public WorksheetObject SelectSheet(string sheetName)
        {
            return this.WorksheetControl.SetCurrent(sheetName);
        }

        /// <summary>
        /// シート名設定
        /// </summary>
        /// <param name="name"></param>
        public void SetSheetName(string name)
        {
            this.WorksheetControl.Current.Sheet.Name = name;
        }

        #endregion

        /// <summary>
        /// 保存
        /// </summary>
        public void Save()
        {
            this.WorksheetControl.Save();
            this.SpreadsheetDoc.Close();
        }

        #region cell

        /// <summary>
        /// セルに書込
        /// </summary>
        /// <param name="column">アルファベット</param>
        /// <param name="row">y座標</param>
        /// <param name="value"></param>
        public void WriteCell(string column, uint row, object value)
        {
            this.WorksheetControl.WriteCell(column, row, value);
        }

        /// <summary>
        /// セルに書込
        /// </summary>
        /// <param name="x">x座標</param>
        /// <param name="y">y座標</param>
        /// <param name="value">値</param>
        public void WriteCell(int x, int y, object value)
        {
            this.WorksheetControl.WriteCell((uint)x, (uint)y, value);
        }

        /// <summary>
        /// セルに書込
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        public void WriteCell(string cell, object value)
        {
            this.WorksheetControl.WriteCell(cell, value);
        }

        /// <summary>
        /// セルの値取得
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <returns></returns>
        public object ReadCell(int x, int y)
        {
            return this.WorksheetControl.ReadCell(x, y);
        }

        public object ReadCell(string cell)
        {
            return this.WorksheetControl.ReadCell(cell);
        }

        #endregion

        /// <summary>
        /// 行コピー&ペースト
        /// </summary>
        /// <param name="from">コピー元となる行数</param>
        /// <param name="to">ペースト先の行数</param>
        /// <param name="isDeep">コピーの深度</param>
        public void RowCopyPaste(uint from, uint to, bool isDeep = true)
        {
            this.WorksheetControl.RowCopyPaste(from, to, isDeep);
        }

        /// <summary>
        /// オブジェクトを書込
        /// </summary>
        /// <param name="dataObject"></param>
        public void WriteDataObject(object dataObject)
        {
            this.WorksheetControl.WriteDataObject(dataObject);
        }

    }
}
