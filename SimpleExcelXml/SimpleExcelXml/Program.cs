using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SimpleExcelXml;

namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {
            Sample sample = new Sample();
            sample.Create();
            sample.CreateTemp();
            Console.ReadLine();
        }
    }

    public class Sample
    {
        /// <summary>
        /// 新しくExcelを作成
        /// </summary>
        public void Create()
        {
            SimpleExcelCreator simpleExcelCreator = new SimpleExcelCreator("operational_test.xlsx");
            // シート追加
            simpleExcelCreator.AddSheet("SampleSheet1");
            simpleExcelCreator.AddSheet("SampleSheet2");
            simpleExcelCreator.AddSheet("SampleSheet3");
            // シート選択
            simpleExcelCreator.SelectSheet("SampleSheet2");
            // 選択中のシート名変更
            simpleExcelCreator.SetSheetName("ChangeName");
            simpleExcelCreator.SelectSheet("SampleSheet1");
            // アルファベットと行番目でデータ書込
            simpleExcelCreator.WriteCell("A", 1, 1000);
            // セル指定でデータ書込
            simpleExcelCreator.WriteCell("A2", "temp");
            simpleExcelCreator.WriteCell("D2", "temp");
            simpleExcelCreator.WriteCell("F2", "temp");
            simpleExcelCreator.WriteCell("E2", "temp");

            // 座標指定でデータ書込
            for (int i = 10; i < 30; i++)
            {
                for (int j = 10; j < 30; j++)
                {
                    simpleExcelCreator.WriteCell(i, j, i.ToString() + j.ToString());
                }
            }
            simpleExcelCreator.WorksheetControl.AddRow(15);
            simpleExcelCreator.SelectSheet("ChangeName");
            // オブジェクトの書込
            simpleExcelCreator.WriteDataObject(this.GetSample());
            simpleExcelCreator.WriteCell(1, 10, "test");
            // 行のコピー&ペースト
            simpleExcelCreator.RowCopyPaste(2, 5, true);
            simpleExcelCreator.RowCopyPaste(5, 7, false);
            simpleExcelCreator.RowCopyPaste(4, 7, false);
            // シートの削除
            simpleExcelCreator.RemoveSheet("SampleSheet3");
            // セルの値読取
            object value1 = simpleExcelCreator.ReadCell(1, 10);
            object value2 = simpleExcelCreator.ReadCell("A2");
            object value3 = simpleExcelCreator.ReadCell("C2");
            simpleExcelCreator.WriteCell(1, 11, DateTime.Now);
            object value4 = simpleExcelCreator.ReadCell(1, 11);
            Console.WriteLine(string.Format("value1 = {0}, value2 = {1}, value3 = {2}", value1, value2, value3));
            // 保存
            simpleExcelCreator.Save();
        }

        /// <summary>
        /// 既存のExcelを元にExcelを作成
        /// </summary>
        public void CreateTemp()
        {
            SimpleExcelCreator simpleExcelCreator = new SimpleExcelCreator("template.xlsx", "use_template.xlsx", true);
            simpleExcelCreator.WriteDataObject(this.GetSample());
            simpleExcelCreator.WriteCell(1, 10, "test");
            simpleExcelCreator.Save();
        }

        /// <summary>
        /// サンプルデータ作成
        /// </summary>
        /// <returns></returns>
        private ExcelSampleObject GetSample()
        {
            ExcelSampleObject excelSampleObject = new ExcelSampleObject
            {
                CompanyName = "Sample会社",
                Name = "黄金色の菓子",
                Price = 10000,
                Introduction = "時代劇等々",
                Remarks = "複数箇所に出力",
                Excluded = "出力しないプロパティ"
            };

            return excelSampleObject;
        }
    }
}
