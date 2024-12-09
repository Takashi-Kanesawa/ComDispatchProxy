using Excel = Microsoft.Office.Interop.Excel;

using (var excelApp = ComDispatchProxy<Excel.Application>.CreateProxy(new Excel.Application()))
using (var wbs = excelApp?.Proxy.Workbooks as ComDispatchProxy<Excel.Workbooks>)
using (var wb = wbs?.Proxy.Add() as ComDispatchProxy<Excel.Workbook>)
using (var wss = wb?.Proxy.Sheets as ComDispatchProxy<Excel.Sheets>)
using (var ws = wss?.Proxy[1] as ComDispatchProxy<Excel.Worksheet>)
using (var targetRange = ws?.Proxy.Range["A1:A10"] as ComDispatchProxy<Excel.Range>)
using (var columns = targetRange?.Proxy.Columns as ComDispatchProxy<Excel.Range>)
using (var windows = excelApp?.Proxy.Windows as ComDispatchProxy<Excel.Windows>)
using (var window = windows?.Proxy[1] as ComDispatchProxy<Excel.Window>)
{
	if (excelApp is not null && targetRange is not null && columns is not null && window is not null)
	{
		excelApp.Proxy.Visible = true;                      // Excelを可視化（見せない方が実は速い）
		window.Proxy.WindowState =
			Excel.XlWindowState.xlMaximized;                // ExcelのWindowを最大化する

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
}

