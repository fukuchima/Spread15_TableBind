using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Spread15_TableBind
{
    internal static class Program
    {
        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            //            var frm = new Form1(); // テーブルの自動連結
            //            var frm = new Form2(); // テーブルの手動連結
            var frm = new Form3(); // テーブルの活用

            var envOS = System.Runtime.InteropServices.RuntimeInformation.OSDescription.ToString();
            var envFW = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription.ToString();
            frm.Text += $" 【 {envOS} : {envFW} 】";

            Application.Run(frm);
        }
    }
}
