using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Spread15_TableBind
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
            this.Text = "テーブルの活用";
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            // データソースの作成
            var ds = new DataSet();

            // Customersテーブルの作成
            var dt1 = ds.Tables.Add("Customers");
            dt1.Columns.AddRange(new DataColumn[] {
                new DataColumn("ID", typeof(string)),
                new DataColumn("Company", typeof(string)),
                new DataColumn("PostalCode", typeof(string)),
                new DataColumn("Address", typeof(string)),
                new DataColumn("PersonInCharge", typeof(string)),
            });
            dt1.Rows.Add("00001", "グレープシティ株式会社", "981-3205", "宮城県仙台市泉区柴山3-1-4", "グレープ太郎");
            dt1.Rows.Add("00002", "オレンジシティ株式会社", "332-0012", "埼玉県川口市本町4-1-8 川口センタービル 3F", "オレンジ花子");
            dt1.AcceptChanges();
            var dv1 = dt1.DefaultView;

            // Invoicesテーブルの作成
            var dt2 = ds.Tables.Add("Invoices");
            dt2.Columns.AddRange(new DataColumn[]
            {
                new DataColumn("ID", typeof(string)),
                new DataColumn("Number", typeof(int)),
                new DataColumn("CustomerID",typeof(string)),
                new DataColumn("品名", typeof(string)),
                new DataColumn("数量", typeof(int)),
                new DataColumn("単価", typeof (int))
            });
            dt2.Rows.Add("20220401A", 1, "00001", "OAモニタ", 2, 29800);
            dt2.Rows.Add("20220401A", 2, "00001", "SDDユニット", 2, 25000);
            dt2.Rows.Add("20220401A", 3, "00001", "出張点検費", 1, 18400);
            dt2.Rows.Add("20220401B", 1, "00002", "OAモニタ", 5, 29800);
            dt2.Rows.Add("20220401B", 2, "00002", "ノートPC", 10, 68000);
            dt2.Rows.Add("20220401B", 3, "00002", "NAS-2T", 2, 175000);
            dt2.Rows.Add("20220401B", 4, "00002", "出張点検費", 1, 18400);
            dt2.AcceptChanges();

            // Invoicesテーブルに[金額]列を追加
            dt2.Columns.Add(new DataColumn("金額", typeof(int)));
            dt2.Columns["金額"].Expression = "[数量] * [単価]";
            dt2.AcceptChanges();
            var dv2 = dt2.DefaultView;

            // ロックの既定値の変更
            fpSpread1.AsWorkbook().Styles[GrapeCity.Spreadsheet.BuiltInStyle.Normal].Locked = false;

            // シートの設定
            var sheet = fpSpread1.Sheets[0];
            sheet.Protect = true;
            sheet.Columns[1, 4].Width = 100;
            sheet.Rows[1].Height = 40;
            sheet.Cells[1, 1].ColumnSpan = 4;
            sheet.Cells[1, 1].Value = "請求書";
            sheet.Cells[1, 1].Border = new FarPoint.Win.LineBorder(Color.Black, 2, false, false, false, true);
            sheet.Cells[1, 1].Font = new Font("メイリオ", 15);
            sheet.Cells[1, 1].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            sheet.Cells[4, 1].ColumnSpan = 4;
            sheet.Cells[6, 2].Value = "様";
            sheet.Cells[8, 1].ColumnSpan = 4;
            sheet.Cells[8, 1].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            sheet.Cells[8, 1].Value = "下記の通り、ご請求申し上げます。";
            sheet.Cells[11, 2].Value = "合計金額";
            sheet.Cells[11, 2].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            sheet.Cells[11, 2, 11, 3].Border = new FarPoint.Win.LineBorder(Color.Black, 1, false, false, false, true);
            sheet.Cells[11, 3].ColumnSpan = 2;
            sheet.Cells[11, 3].Font = sheet.Cells[1, 1].Font;
            sheet.Cells[11, 3].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right;

            // セルのデータ連結
            var data1 = new FarPoint.Win.Spread.Data.SpreadDataBindingAdapter();
            data1.Spread = fpSpread1;
            data1.SheetName = "Sheet1";
            data1.MapperInfo = new FarPoint.Win.Spread.Data.MapperInfo(3, 1, 1, 1); // PostalCode

            var data2 = new FarPoint.Win.Spread.Data.SpreadDataBindingAdapter();
            data2.Spread = fpSpread1;
            data2.SheetName = "Sheet1";
            data2.MapperInfo = new FarPoint.Win.Spread.Data.MapperInfo(4, 1, 1, 1); // Address

            var data3 = new FarPoint.Win.Spread.Data.SpreadDataBindingAdapter();
            data3.Spread = fpSpread1;
            data3.SheetName = "Sheet1";
            data3.MapperInfo = new FarPoint.Win.Spread.Data.MapperInfo(6, 1, 1, 1); // PersonInCharge

            // テーブルの作成と設定
            var anchorRow = 14;
            var iTable = sheet.AsWorksheet().Range($"B{anchorRow}:E{anchorRow + dv2.Table.Rows.Count}").CreateTable(true);
            var rowStart = iTable.Range.Row + 1;
            var rowEnd = iTable.Range.Row2 + 1;
            iTable.Name = "table";
            iTable.AutoGenerateColumns = false;
            iTable.TableColumns[0].DataField = "品名";
            iTable.TableColumns[1].DataField = "数量";
            iTable.TableColumns[2].DataField = "単価";
            iTable.TableColumns[3].DataField = "金額";
            iTable.DataBodyRange.Columns[2].NumberFormat = "\\#,##0";
            iTable.DataBodyRange.Columns[3].NumberFormat = "\\#,##0";
            iTable.DataBodyRange.Columns[3].Locked = true;
            iTable.ShowTotals = true;
            iTable.TableColumns[0].Total.Value = "合計";
            iTable.TableColumns[3].Total.Formula = $"TEXT(SUM(E{rowStart}:E{rowEnd}),\"\\#,##0\")";

            // 合計行のロック
            var table = sheet.GetTable("table");
            table.TableTotalRowStyle = new FarPoint.Win.Spread.StyleInfo() { Locked = true, HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right };

            // 合計金額セルの数式設定
            sheet.Cells[11, 3].Formula = iTable.TableColumns[3].Total.ToString();

            // コンボボックスの設定
            var dv = dt2.DefaultView;
            var dt = dv.ToTable(true, new string[] { "ID" });
            for(var i = 0; i < dt.Rows.Count; i++)
            {
                comboBox1.Items.Add(dt.Rows[i][0].ToString());
            }
            comboBox1.SelectedIndexChanged += (s, ea) =>
            {
                // テーブルのデータ更新
                dv2.RowFilter = $"ID='{comboBox1.Text}'";
                iTable.DataSource = dv2;

                // 連結セルのデータ更新
                dv1.RowFilter = $"ID='{dv2[0][2]}'";
                data1.DataSource = dv1.ToTable(false, new string[] { "PostalCode" });
                data1.FillSpreadDataByDataSource();
                data2.DataSource = dv1.ToTable(false, new string[] { "Address" });
                data2.FillSpreadDataByDataSource();
                data3.DataSource = dv1.ToTable(false, new string[] { "PersonInCharge" });
                data3.FillSpreadDataByDataSource();
            };
            comboBox1.SelectedIndex = 0;
        }
    }
}
