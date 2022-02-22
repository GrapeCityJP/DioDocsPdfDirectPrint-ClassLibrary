using DioDocsPdfPrintLibrary;
using GrapeCity.Documents.Common;
using GrapeCity.Documents.Pdf;
using System.Drawing.Printing;

Console.WriteLine("PDFファイルをプリンタダイアログ表示なしで直接印刷します");

// トライアル版または製品版のライセンスキーを設定しない場合はPDFファイルから読み込めるページ数が
// 5ページに制限されます。そのため印刷できるページ数も5ページまでになります
// https://docs.grapecity.com/help/diodocs/pdf/#licenseinfo.html
GcPdfDocument.SetLicenseKey("");

// PDFファイルを読み込み
var fs = File.OpenRead("diodocs_startguide_excel_template.pdf");
var doc = new GcPdfDocument();
doc.Load(fs);

// PDFファイルを印刷
GcPdfPrintManager pm = new GcPdfPrintManager();
pm.Doc = doc;
pm.PrinterSettings = new PrinterSettings();
pm.PrinterSettings.PrinterName = "Microsoft Print to PDF";
pm.OutputRange = new OutputRange(5, 10);
pm.PageScaling = PageScaling.FitToPrintableArea;
pm.Print();