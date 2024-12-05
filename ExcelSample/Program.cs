using Excel = Microsoft.Office.Interop.Excel;

var excelApp = ComDispatchProxy<Excel.Application>.CreateProxy(new Excel.Application());
var wbs = ComDispatchProxy<Excel.Workbooks>.CreateProxy(excelApp?.Workbooks);
var wb = ComDispatchProxy<Excel.Workbook>.CreateProxy(wbs?.Add());
var wss = ComDispatchProxy<Excel.Sheets>.CreateProxy(wb?.Sheets);
var ws = ComDispatchProxy<Excel.Worksheet>.CreateProxy(wss?[1] as Excel.Worksheet);
var targetRange = ComDispatchProxy<Excel.Range>.CreateProxy(ws?.Range["A1:A10"]);
var columns = ComDispatchProxy<Excel.Range>.CreateProxy(targetRange?.Columns);

if (excelApp is not null && targetRange is not null && columns is not null)
{
    excelApp.Visible = true;                      // Excelを可視化（見せない方が実は速い）
    var startDate = DateTime.Today;               // 今日から10日分の日付データ作成
    var dates = new DateTime[10];
    for (int i = 0; i < 10; i++)
    {
        dates[i] = startDate.AddDays(i);
    }

    targetRange.NumberFormat = "yyyy年mm月dd日";  // 編集範囲の書式指定
    targetRange.Value = dates;                    // 編集範囲にまとめて入力（一個ずつやるとめちゃ遅い）
    columns.AutoFit();                // 列幅を自動調整
}

