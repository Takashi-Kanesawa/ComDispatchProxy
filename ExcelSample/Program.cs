using System.Runtime.InteropServices;

using Excel = Microsoft.Office.Interop.Excel;
var excelApp = new Excel.Application();                 // Excel起動
var wbs = excelApp?.Workbooks;
var wb = wbs.Add();                                     // Workbook取得
var wss = wb?.Sheets;
var ws = wss[1] as Excel.Worksheet;                     // Worksheet取得
var targetRange = ws?.Range["A1:A10"] as Excel.Range;   // 編集範囲のRangeオブジェクトを取得
var columns = targetRange?.Columns as Excel.Range;
if (excelApp is not null && wb is not null && ws is not null && targetRange is not null)
{
    try
    {
        excelApp.Visible = true;                        // Excelを可視化（見せない方が実は速い）
        var startDate = DateTime.Today;                 // 今日から10日分の日付データ作成
        var dates = new DateTime[10];
        for (int i = 0; i < 10; i++)
        {
            dates[i] = startDate.AddDays(i);
        }
        targetRange.NumberFormat = "yyyy年mm月dd日";    // 編集範囲の書式指定
        targetRange.Value = dates;                      // 編集範囲にまとめて入力
        columns.AutoFit();                              // 列幅を自動調整
    }
    finally
    {
        // プラットフォーム固有のメソッド(Marshal.ReleaseComObject)を使うと警告が出るので抑止
#pragma warning disable CA1416
        Marshal.ReleaseComObject(columns);             // COMオブジェクトの解放
        Marshal.ReleaseComObject(targetRange);
        Marshal.ReleaseComObject(ws);
        Marshal.ReleaseComObject(wss);
        Marshal.ReleaseComObject(wb);
        Marshal.ReleaseComObject(wbs);
        Marshal.ReleaseComObject(excelApp);
#pragma warning restore CA1416
    }
}