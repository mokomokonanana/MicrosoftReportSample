# MicrosoftReportSample

## 「Microsoft RDLC レポートデザイナー」インストール

1. 拡張機能 > 拡張機能の管理 > Microsoft RDLC レポートデザイナー > ダウンロード
![001](https://user-images.githubusercontent.com/58506797/72318717-5bdce980-36e0-11ea-935b-6e0e5473e588.png)
![002](https://user-images.githubusercontent.com/58506797/72318877-dc034f00-36e0-11ea-8f66-4d091edc44e1.png)
![003](https://user-images.githubusercontent.com/58506797/72318994-343a5100-36e1-11ea-8e9e-944b98734cd1.png)

1. Visual Studio 終了

1. 拡張機能インストール
![004](https://user-images.githubusercontent.com/58506797/72319079-88ddcc00-36e1-11ea-9bdd-f2dd7d840127.png)
![005](https://user-images.githubusercontent.com/58506797/72319118-a9a62180-36e1-11ea-8a45-fb99c424047a.png)
![006](https://user-images.githubusercontent.com/58506797/72319117-a90d8b00-36e1-11ea-8207-8bbf88b91af2.png)

## 「Microsoft.ReportingServices.ReportViewerControl.Winforms」インストール
1. プロジェクト(右クリック) > NuGetパッケージ管理 > Microsoft.ReportingServices.ReportViewerControl.Winforms > インストール
![007](https://user-images.githubusercontent.com/58506797/72319902-c9d6e000-36e3-11ea-8df2-3c5c85b4b048.png)
![008](https://user-images.githubusercontent.com/58506797/72319907-cc393a00-36e3-11ea-976a-ac7306a55918.png)
![009](https://user-images.githubusercontent.com/58506797/72320888-ebd16200-36e5-11ea-8a38-a9c32919dfc1.png)
![010](https://user-images.githubusercontent.com/58506797/72320891-ee33bc00-36e5-11ea-8a18-0f3cd5f10224.png)

## レポートの作成
1. プロジェクト(右クリック) > 追加 > 新しい項目 > レポート
![011](https://user-images.githubusercontent.com/58506797/72322972-f93d1b00-36ea-11ea-89cf-f1dc9c21a4b1.png)
![012](https://user-images.githubusercontent.com/58506797/72322974-f93d1b00-36ea-11ea-8d41-1b1fb502df89.png)
2. レポートの中身設定（ここではとりあえずで一覧を表示）
![013](https://user-images.githubusercontent.com/58506797/72322975-f9d5b180-36ea-11ea-862a-ebe58a6f5cf3.png)
![014](https://user-images.githubusercontent.com/58506797/72323506-42419f00-36ec-11ea-8b7a-cdc40216fe06.png)
![015](https://user-images.githubusercontent.com/58506797/72322977-f9d5b180-36ea-11ea-86a2-4bfa260a096b.png)
![016](https://user-images.githubusercontent.com/58506797/72322971-f93d1b00-36ea-11ea-9435-74f5e6ec32dd.png)

## ReportViewを表示
1. ツールボックス > Microsoft SQL Server > ReportViewer(ドラッグ＆ドロップ)
![017](https://user-images.githubusercontent.com/58506797/72323862-09ee9080-36ed-11ea-9007-b3d00f947ef5.png)
2. レポート設定
![018](https://user-images.githubusercontent.com/58506797/72323863-0a872700-36ed-11ea-9a7c-19c58d99c297.png)
![019](https://user-images.githubusercontent.com/58506797/72323861-09ee9080-36ed-11ea-8a7c-71bce22e15fe.png)

## Reportにデータ設定

```C#
private void Form1_Load(object sender, EventArgs e)
{
  this.DataSet1.DataTable1.AddDataTable1Row("test1", "test1-1", 100);
  this.DataSet1.DataTable1.AddDataTable1Row("test1", "test1-2", 100);
  this.DataSet1.DataTable1.AddDataTable1Row("test2", "test2-1", 100);
  this.DataSet1.DataTable1.AddDataTable1Row("test2", "test2-2", 100);
  this.reportViewer1.RefreshReport();
}
```
![020](https://user-images.githubusercontent.com/58506797/72324288-ec6df680-36ed-11ea-8505-ef79b8564787.png)

## 画面に表示せずにエクセルとして出力
```C#
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
```
