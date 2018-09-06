using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace SimpleExcelXml
{

    /// <summary>
    /// ワークシートオブジェクト
    /// </summary>
    public class WorksheetObject
    {
        public WorksheetPart WorksheetPart { get; set; }

        public SheetData SheetData { get; set; }

        public Sheet Sheet { get; set; }

        public WorksheetObject(WorksheetPart worksheetPart, Sheet sheet)
        {
            this.WorksheetPart = worksheetPart;
            this.SheetData = (SheetData)this.WorksheetPart.Worksheet.Elements<SheetData>().First();
            this.Sheet = sheet;
        }
    }

    /// <summary>
    /// ワークシート操作クラス
    /// </summary>
    public class WorksheetControl
    {
        #region property

        /// <summary>
        /// 操作中のオブジェクト
        /// </summary>
        public WorksheetObject Current { get; set; }

        public List<WorksheetObject> WorksheetObjects { get; set; }

        #endregion

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="workbookPart"></param>
        public WorksheetControl(WorkbookPart workbookPart)
        {
            this.WorksheetObjects = new List<WorksheetObject>();
            this.Init(workbookPart);
        }

        /// <summary>
        /// 初期化
        /// </summary>
        /// <param name="workbookPart"></param>
        private void Init(WorkbookPart workbookPart)
        {
            foreach (WorksheetPart worksheetPart in workbookPart.WorksheetParts)
            {
                Sheet sheet = null;
                foreach (Sheet s in workbookPart.Workbook.Sheets)
                {
                    string id = workbookPart.GetIdOfPart(worksheetPart);
                    if (s.Id.Value.Equals(id))
                    {
                        sheet = s;
                    }
                }
                this.WorksheetObjects.Add(new WorksheetObject(worksheetPart, sheet));
            }

            this.Current = this.WorksheetObjects.FirstOrDefault();
        }

        #region sheet

        /// <summary>
        /// 操作シート設定
        /// </summary>
        /// <param name="index"></param>
        public WorksheetObject SetCurrent(int index)
        {
            this.Current = this.WorksheetObjects[index];
            return this.Current;
        }

        /// <summary>
        /// 操作シート設定
        /// </summary>
        /// <param name="name"></param>
        public WorksheetObject SetCurrent(string name)
        {
            this.Current = this.WorksheetObjects.First(w => w.Sheet.Name.Value.Equals(name));
            return this.Current;
        }

        /// <summary>
        /// 操作シート追加(管理するオブジェクトを追加) 
        /// </summary>
        /// <param name="worksheetObject"></param>
        public void AddWorksheetObject(WorksheetObject worksheetObject)
        {
            this.WorksheetObjects.Add(worksheetObject);
            if (this.WorksheetObjects.Count.Equals(1))
            {
                this.Current = worksheetObject;
            }
        }

        /// <summary>
        /// ワークシートオブジェクトを削除
        /// </summary>
        /// <param name="openXmlPart"></param>
        public void RemoveWorksheetObject(OpenXmlPart openXmlPart)
        {
            var worksheetObject = this.WorksheetObjects.FirstOrDefault(w => w.WorksheetPart.Equals(openXmlPart));
            if (worksheetObject != null)
            {
                this.WorksheetObjects.Remove(worksheetObject);
            }
        }

        #endregion

        #region writecell

        /// <summary>
        /// セルに書込
        /// </summary>
        /// <param name="column">アルファベット</param>
        /// <param name="y">y座標</param>
        /// <param name="value"></param>
        public void WriteCell(string column, uint y, object value)
        {
            y++;
            Row row = this.GetRowElement(y);
            Cell cell = this.GetCellElement(row, y, column);
            cell.CellValue = new CellValue(value.ToString());
            cell.DataType = new EnumValue<CellValues>(value.GetType().GetCellValuesType());
        }

        /// <summary>
        /// セルに書込
        /// </summary>
        /// <param name="x">x座標</param>
        /// <param name="y">y座標</param>
        /// <param name="value">値</param>
        public void WriteCell(int x, int y, object value)
        {
            x++;
            string column = this.GetColumnFromIndex(x);
            this.WriteCell(column, (uint)y, value);
        }

        /// <summary>
        /// セルに書込
        /// </summary>
        /// <param name="cell">セル座標(例：A1)</param>
        /// <param name="value">値</param>
        public void WriteCell(string cell, object value)
        {
            string column = Regex.Replace(cell, @"[^A-Z]", "");
            int y = int.Parse(Regex.Replace(cell, @"[^0-9]", ""));
            y--;
            this.WriteCell(column, (uint)y, value);

        }

        /// <summary>
        /// 数値から座標アルファベットに変換
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        private string GetColumnFromIndex(int index)
        {
            // 1桁目 + 2桁目(十の位) * 26 = 2桁目の計算
            // 1桁目 + 2桁目(十の位) * 26 + 3桁目(百の位) * 676
            string column;
            int n1 = index / 676;
            int n2 = index % 676 / 26;
            int n3 = index % 676 % 26 + 1;
            column = n1 > 0 ? ((char)(n1 + 64)).ToString() : "";
            column += n2 > 0 ? ((char)(n2 + 64)).ToString() : "";
            column += ((char)(n3 + 64)).ToString();
            return column;
        }

        /// <summary>
        /// 列をインデックスに変換
        /// </summary>
        /// <param name="column"></param>
        /// <returns></returns>
        private uint GetIndexFromColumn(string column)
        {
            column = Regex.Replace(column, @"[^A-Z]", "");
            Queue<char> chars = new Queue<char>(column.Reverse());
            int i = 0;
            uint index = 0;
            while (chars.Count > 0)
            {
                char c = chars.Dequeue();
                uint no = ((uint)c) - 64;
                if (i == 2)
                    index += no * 676;
                else if (i == 1)
                    index += no * 26;
                else
                    index += no;

                i++;
            }

            return index;
        }

        #endregion

        #region copy&paste

        /// <summary>
        /// 行コピー＆ペースト
        /// </summary>
        /// <param name="from"></param>
        /// <param name="to"></param>
        /// <param name="isInsert">行挿入かどうか</param>
        /// <param name="isDeep"></param>
        public void RowCopyPaste(uint from, uint to, bool isInsert, bool isDeep = true)
        {
            Row clone = this.CloneRow(from, to, isDeep);
            this.AddRow(clone, isInsert);
        }

        /// <summary>
        /// 行コピー
        /// </summary>
        /// <param name="from"></param>
        /// <param name="to"></param>
        /// <returns></returns>
        private Row CloneRow(uint from, uint to, bool isDeep = true)
        {
            Row fromRow = this.GetRowElement(from);
            Row clone = (Row)fromRow.CloneNode(isDeep);
            clone.RowIndex = to;
            this.ReplaceCellReference(clone);
            return clone;
        }

        /// <summary>
        /// セルの参照行数を上書
        /// </summary>
        /// <param name="row"></param>
        /// <param name="index"></param>
        private void ReplaceCellReference(Row row)
        {
            row.Elements<Cell>().ForEach(c =>
            {
                string reference = Regex.Replace(c.CellReference.Value, @"[^A-Z]", "");
                c.CellReference = reference + row.RowIndex.Value.ToString();
            });
        }

        #endregion

        #region addrow,cell

        /// <summary>
        /// 行挿入
        /// </summary>
        /// <param name="index"></param>
        public void AddRow(uint index)
        {
            Row row = new Row { RowIndex = index };
            this.AddRow(row, true);
        }

        /// <summary>
        /// 行追加
        /// </summary>
        /// <param name="row"></param>
        private void AddRow(Row row, bool isInsert)
        {
            IEnumerable<Row> overRows = null;
            Row nearRow = null;
            if (isInsert)
            {
                overRows = this.Current.SheetData.Elements<Row>().Where(r => r.RowIndex.Value >= row.RowIndex.Value);
                nearRow = overRows.FirstOrDefault();
            }
            else
            {
                nearRow = this.Current.SheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex.Value >= row.RowIndex.Value);
            }

            if (isInsert)
            {
                overRows.ForEach(r =>
                {
                    r.RowIndex.Value++;
                    this.ReplaceCellReference(r);
                });
            }

            if (nearRow != null)
            {
                nearRow.InsertBeforeSelf<Row>(row);
                if (nearRow.RowIndex.Value.Equals(row.RowIndex.Value))
                    nearRow.Remove();
            }
            else
            {
                this.Current.SheetData.Append(row);
            }
        }

        /// <summary>
        /// セル追加
        /// </summary>
        /// <param name="row"></param>
        /// <param name="cell"></param>
        private void AddCell(Row row, Cell cell, bool isInsert)
        {
            uint columnIndex = this.GetIndexFromColumn(cell.CellReference.Value);
            IEnumerable<Cell> overCells = null;
            Cell nearCell = null;
            if (isInsert)
            {
                overCells = row.Elements<Cell>().Where(c => columnIndex <= this.GetIndexFromColumn(c.CellReference.Value));
                nearCell = overCells.FirstOrDefault();
            }
            else
            {
                nearCell = row.Elements<Cell>().FirstOrDefault(c => columnIndex <= this.GetIndexFromColumn(c.CellReference.Value));
            }

            if (isInsert)
            {
                overCells.ForEach(c =>
                {
                    uint index = this.GetIndexFromColumn(c.CellReference.Value);
                    index++;
                    c.CellReference.Value = this.GetColumnFromIndex((int)index) + row.RowIndex.Value.ToString();
                });
            }

            if (nearCell != null)
                nearCell.InsertBeforeSelf<Cell>(cell);
            else
                row.Append(cell);
        }

        #endregion

        /// <summary>
        /// 行要素取得。
        /// </summary>
        /// <param name="xml"></param>
        /// <param name="index"></param>
        /// <returns></returns>
        private Row GetRowElement(uint index)
        {
            Row row = this.Current.SheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex.Value.Equals(index));
            if (row == null)
            {
                row = new Row() { RowIndex = index };
                this.AddRow(row, false);
            }

            return row;
        }

        /// <summary>
        /// セル要素取得
        /// </summary>
        /// <param name="row"></param>
        /// <param name="index"></param>
        /// <param name="column"></param>
        /// <returns></returns>
        private Cell GetCellElement(Row row, uint index, string column)
        {
            string cellRef = column + index.ToString();
            Cell cell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference.Value.Equals(cellRef));
            if (cell == null)
            {
                cell = new Cell() { CellReference = cellRef };
                this.AddCell(row, cell, false);
            }

            return cell;
        }

        /// <summary>
        /// 保存
        /// </summary>
        public void Save()
        {
            this.WorksheetObjects.ForEach(w => w.WorksheetPart.Worksheet.Save());
        }

        #region writedataobject

        /// <summary>
        /// クラスを基にレポート内容作成
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dataObject"></param>
        public void WriteDataObject(object dataObject)
        {
            PropertyInfo[] propertyInfos = dataObject.GetType().GetProperties();
            foreach (PropertyInfo propertyInfo in propertyInfos)
            {
                CellAttribute cellAttribute = propertyInfo.GetCustomAttribute<CellAttribute>();
                if (cellAttribute != null)
                {
                    object value = propertyInfo.GetGetMethod().Invoke(dataObject, null);
                    if (value != null)
                        this.WriteCell(cellAttribute, value);
                }
            }
        }

        /// <summary>
        /// 書込み先の情報を基にエクセル内容作成
        /// </summary>
        /// <param name="cellAttribute"></param>
        /// <param name="value"></param>
        private void WriteCell(CellAttribute cellAttribute, object value)
        {
            string format = cellAttribute.Format;
            foreach (string position in cellAttribute.Positions)
            {
                if (string.IsNullOrEmpty(cellAttribute.Format))
                {
                    this.WriteCell(position, value);
                }
                else
                {
                    this.WriteCell(position, string.Format(cellAttribute.Format, value));
                }
            }
        }

        #endregion

    }

}
