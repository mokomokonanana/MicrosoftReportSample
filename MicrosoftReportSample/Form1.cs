using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MicrosoftReportSample
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.DataSet1.DataTable1.AddDataTable1Row("test1", "test1-1", 100);
            this.DataSet1.DataTable1.AddDataTable1Row("test1", "test1-2", 100);
            this.DataSet1.DataTable1.AddDataTable1Row("test2", "test2-1", 100);
            this.DataSet1.DataTable1.AddDataTable1Row("test2", "test2-2", 100);
            this.reportViewer1.RefreshReport();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // ReportViewer作成（画面に表示しない）
            var rv = new Microsoft.Reporting.WinForms.ReportViewer();
            //  「MicrosoftReportSample.Report1.rdlc」は「プロジェクト名.Reportテンプレートファイル名」
            rv.LocalReport.ReportEmbeddedResource = "MicrosoftReportSample.Report1.rdlc";

            // データ作成
            var ds = new DataSet1();
            ds.DataTable1.AddDataTable1Row("test1", "test1-1", 100);
            ds.DataTable1.AddDataTable1Row("test1", "test1-2", 100);
            ds.DataTable1.AddDataTable1Row("test2", "test2-1", 100);
            ds.DataTable1.AddDataTable1Row("test2", "test2-2", 100);

            // データとReportの紐づけ
            rv.LocalReport.DataSources.Add(
                new Microsoft.Reporting.WinForms.ReportDataSource("DataSet1", ds.DataTable1.Copy()));
                // 「DataSet1」はレポート作成時覚えておいた名前

            // Reportにデータ表示
            rv.RefreshReport();

            // エクセル出力
            var bytes = rv.LocalReport.Render("EXCELOPENXML");
            using (var fs = new System.IO.FileStream("test.xlsx", System.IO.FileMode.Create)) // exeと同階層にファイルを出力
            {
                fs.Write(bytes, 0, bytes.Length);
            }
        }
    }
}
