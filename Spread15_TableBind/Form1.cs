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
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            this.Text = "テーブルの自動連結";
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // データソースの作成
            var dt = new DataTable();
            dt.Columns.AddRange(new DataColumn[]
            {
                new DataColumn("品名", typeof(string)),
                new DataColumn("数量", typeof(int)),
                new DataColumn("単価", typeof (int))
            });
            dt.Rows.Add("OAモニタ", 2, 29800);
            dt.Rows.Add("SDDユニット", 2, 25000);
            dt.Rows.Add("出張点検費", 1, 18400);
            dt.AcceptChanges();

            // データソースの変更
            dt.Columns.Add(new DataColumn("金額", typeof(int)));
            dt.Columns["金額"].Expression = "[数量] * [単価]";
            dt.AcceptChanges();

            // シートの設定
            var sheet = fpSpread1.Sheets[0];
            sheet.SetColumnWidth(1, 100);

            // テーブルの作成と設定
            var iTable = sheet.AsWorksheet().Range("B2:E5").CreateTable(true);
            iTable.AutoGenerateColumns = true;
            iTable.DataSource = dt;
        }
    }
}
