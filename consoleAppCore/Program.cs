using System;
using System.Data;
using ClosedXML.Excel;

namespace consoleAppCore
{
    class Program
    {
        static void Main(string[] args)
        {
            //ファイルパスの指定
            var path = "[エクセルファイルの指定]";

            //エクセルを開く
            using (var wb = new XLWorkbook(path))
            {
                //シートの指定
                var sheet = wb.Worksheet("Sheet1");

                //テーブルを設定
                var table = wb.Table("Table1");



                //A1セルの値を取得
                var cellA1value = sheet.Cell("A1").Value;

                //アドレスを数字でも指定可能
                var cellA3value = sheet.Cell(3,1).Value;
                Console.WriteLine(cellA3value);



                //A1セルに値を設定
                sheet.Cell("A1").SetValue("test");

                //アドレスを数字でも指定可能
                sheet.Cell(1, 1).SetValue("test");



                //A3セルのハイパーリンクを取得
                var hyperlink = sheet.Cell("A3").Hyperlink.ExternalAddress;



                //テーブル内の値を列挙する
                foreach (var row in table.DataRange.Rows())
                {
                    //現在の行の全セルを取得
                    var cells = row.Cells();

                    //それぞれのセルに記載されている値を表示
                    foreach (var cell in cells)
                    {
                        Console.Write(cell.Value);
                        Console.Write(" ");
                    }

                    Console.WriteLine();
                }



                //データテーブル型の変数を作成
                var newData = new DataTable();
                newData.Columns.Add("Comumn1", typeof(string));
                newData.Columns.Add("Comumn2", typeof(string));
                newData.Columns.Add("Comumn3", typeof(string));

                newData.Rows.Add("111", "222", "333");

                //テーブルに行を追加
                table.AppendData(newData);



                //エクセルファイルを保存する
                wb.Save();
            }
        }
    }
}
