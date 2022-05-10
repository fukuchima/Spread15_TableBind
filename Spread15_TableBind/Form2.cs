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
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
            this.Text = "テーブルの手動連結";
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            // データソースの作成
            var dt = new DataTable();
            dt.Columns.AddRange(new DataColumn[]
            {
                new DataColumn("品名", typeof(string)),
                new DataColumn("数量", typeof(int)),
                new DataColumn("単価", typeof (int)),
            });
            dt.Rows.Add("OAモニタ", 2, 29800);
            dt.Rows.Add("SDDユニット", 2, 25000);
            dt.Rows.Add("出張点検費", 1, 18400);
            dt.AcceptChanges();

            // シートの設定
            var sheet = fpSpread1.Sheets[0];
            sheet.Columns[1, 4].Width = 100;

            // テーブルの作成と設定
            var iTable = sheet.AsWorksheet().Range("B2:E5").CreateTable(true);
            iTable.Name = "table";
            iTable.AutoGenerateColumns = false;
            iTable.TableColumns[0].DataField = "品名";
            iTable.TableColumns[1].DataField = "数量";
            iTable.TableColumns[2].DataField = "単価";
            iTable.TableColumns[3].Header.Value = "金額";
            iTable.DataBodyRange.Columns[2].NumberFormat = "\\#,##0";
            iTable.DataBodyRange.Columns[3].NumberFormat = "\\#,##0";
            iTable.ShowTotals = true;
            iTable.TableColumns[0].Total.Value = "合計";
            iTable.TableColumns[3].Total.Formula = "SUM([金額])";
            iTable.TableColumns[3].Total.NumberFormat = "\\#,##0";
            iTable.DataSource = dt;

            var table = sheet.GetTable("table");
            table.SetDataColumnFormula(3, "[@単価]*[@数量]");
        }
    }
}
