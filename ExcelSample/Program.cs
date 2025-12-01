using Excel = Microsoft.Office.Interop.Excel;
using ComDispatchProxy;
using ComDispatchProxy.ProxyFactories;


var excelFactory = new InteropExcelProxyFactory("ExcelInterfaces.xml");

// Excel アプリケーションのインスタンスを作成し、ComDispatchProxy を介して管理する
using (var excelAppRoot = ComDispatchProxy<Excel.Application>.CreateProxy(excelFactory, new Excel.Application()))
{
    var excelApp = excelAppRoot.Proxy;

    // 個別のオブジェクトは、暗黙的に呼び出されるCOMインターフェイスもあるため、Proxy経由で取得する
    var wbs = excelApp.Workbooks;
    var wb = wbs?.Add();
    var wss = wb?.Worksheets;
    var ws = wss?[1] as Excel.Worksheet; // Excel.Sheetsのインデクサはobject型のため、明示的にExcel.Worksheetにする
    var targetRange = ws?.Range["A1:A10"];
    var columns = targetRange?.Columns;
    var windows = excelApp?.Windows;
    var window = windows?[1];
    {
        // 全ての必要なオブジェクトが生成されている場合に処理を実行
        if (excelApp is not null && targetRange is not null && columns is not null && window is not null)
        {
            // Excel を可視化する（非表示のままだと高速だが、デバッグ時に表示する方が便利）
            excelApp.Visible = true;

            // Excel ウィンドウを最大化する
            window.WindowState = Excel.XlWindowState.xlMaximized;

            // 今日から始まる10日間の日付データを準備（入力テスト用）
            var startDate = DateTime.Today;
            var dates = new DateTime[10];
            for (int i = 0; i < 10; i++)
            {
                dates[i] = startDate.AddDays(i);
            }

            // 範囲 "A1:A10" のセル書式を「yyyy年mm月dd日」に設定
            targetRange.NumberFormat = "yyyy年mm月dd日";

            // 範囲 "A1:A10" に日付データを一括入力（1セルずつより効率的）
            targetRange.Value = dates;

            // 列幅を自動調整して内容に合わせる
            columns.AutoFit();
        }
    }
}
