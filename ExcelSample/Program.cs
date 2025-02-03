using Excel = Microsoft.Office.Interop.Excel;

DispatchProxyFactory.InitializeFromXml("ExcelInterfaces.xml");

// Excel アプリケーションのインスタンスを作成し、ComDispatchProxy を介して管理する
using (var excelApp = ComDispatchProxy<Excel.Application>.CreateProxy(new Excel.Application()))

// 個別のオブジェクトは、暗黙的に呼び出されるCOMインターフェイスもあるため、Proxy経由で取得する
using (var wbs = excelApp.Proxy.Workbooks as ComDispatchProxy<Excel.Workbooks>)
using (var wb = wbs?.Proxy.Add() as ComDispatchProxy<Excel.Workbook>)
using (var wss = wb?.Proxy.Worksheets as ComDispatchProxy<Excel.Sheets>)
using (var ws = wss?.Proxy[1] as ComDispatchProxy<Excel.Worksheet>)
using (var targetRange = ws?.Proxy.Range["A1:A10"] as ComDispatchProxy<Excel.Range>)
using (var columns = targetRange?.Proxy.Columns as ComDispatchProxy<Excel.Range>)
using (var windows = excelApp?.Proxy.Windows as ComDispatchProxy<Excel.Windows>)
using (var window = windows?.Proxy[1] as ComDispatchProxy<Excel.Window>)
{
    // 全ての必要なオブジェクトが生成されている場合に処理を実行
    if (excelApp is not null && targetRange is not null && columns is not null && window is not null)
    {
        // Excel を可視化する（非表示のままだと高速だが、デバッグ時に表示する方が便利）
        excelApp.Proxy.Visible = true;

        // Excel ウィンドウを最大化する
        window.Proxy.WindowState = Excel.XlWindowState.xlMaximized;

        // 今日から始まる10日間の日付データを準備（入力テスト用）
        var startDate = DateTime.Today;
        var dates = new DateTime[10];
        for (int i = 0; i < 10; i++)
        {
            dates[i] = startDate.AddDays(i);
        }

        // 範囲 "A1:A10" のセル書式を「yyyy年mm月dd日」に設定
        targetRange.Proxy.NumberFormat = "yyyy年mm月dd日";

        // 範囲 "A1:A10" に日付データを一括入力（1セルずつより効率的）
        targetRange.Proxy.Value = dates;

        // 列幅を自動調整して内容に合わせる
        columns.Proxy.AutoFit();
    }
}
