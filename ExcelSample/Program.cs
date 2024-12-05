using Excel = Microsoft.Office.Interop.Excel;

using (var excelApp = ComDispatchProxy<Excel.Application>.CreateProxy(new Excel.Application()))
using (var wbs = ComDispatchProxy<Excel.Workbooks>.CreateProxy(excelApp?.Proxy.Workbooks))
using (var wb = ComDispatchProxy<Excel.Workbook>.CreateProxy(wbs?.Proxy.Add()))
using (var wss = ComDispatchProxy<Excel.Sheets>.CreateProxy(wb?.Proxy.Sheets))
using (var ws = ComDispatchProxy<Excel.Worksheet>.CreateProxy(wss?.Proxy[1] as Excel.Worksheet))
using (var targetRange = ComDispatchProxy<Excel.Range>.CreateProxy(ws?.Proxy.Range["A1:A10"]))
using (var columns = ComDispatchProxy<Excel.Range>.CreateProxy(targetRange?.Proxy.Columns))

if (excelApp is not null && targetRange is not null && columns is not null)
{
    excelApp.Proxy.Visible = true;                      // Excelを可視化（見せない方が実は速い）
    var startDate = DateTime.Today;                     // 今日から10日分の日付データ作成
    var dates = new DateTime[10];
    for (int i = 0; i < 10; i++)
    {
        dates[i] = startDate.AddDays(i);
    }

    targetRange.Proxy.NumberFormat = "yyyy年mm月dd日";  // 編集範囲の書式指定
    targetRange.Proxy.Value = dates;                    // 編集範囲にまとめて入力（一個ずつやるとめちゃ遅い）
    columns.Proxy.AutoFit();                            // 列幅を自動調整
}

