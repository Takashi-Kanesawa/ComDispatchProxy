using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using ComDispatchProxy;
using ComDispatchProxy.ProxyFactories;
using System.ComponentModel.Design.Serialization;

Native.RunOnce(aggressive: true);
Console.WriteLine();
Native.RunOnce(aggressive: false);
GC.Collect();
GC.WaitForPendingFinalizers();
GC.Collect();


static class Native
{
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    static int TryGetExcelPid(Excel.Application app)
    {
        // Excel.Application.Hwnd は int
        var hwnd = new IntPtr(app.Hwnd);
        Native.GetWindowThreadProcessId(hwnd, out var pid);
        return unchecked((int)pid);
    }

    public static void RunOnce(bool aggressive)
    {
        ComProxyLogConfig.Enabled = true;
        ComProxyLogConfig.Writer = Console.WriteLine;

        ComProxyConfig.AggressiveReleaseComObjects = aggressive;

        var excelFactory = new InteropExcelProxyFactory(typeof(Excel.Application).Assembly.GetTypes());

        int pid = -1;

        {
            // ここをメソッド内に閉じ込めるのが重要（ローカル参照が外に残らない）
            using (var root = ComDispatchProxy<Excel.Application>.CreateProxy(
                    excelFactory,
                    new Excel.Application()))
            {
                var app = root.Proxy;

                pid = TryGetExcelPid(app);
                Console.WriteLine($"[TEST] AggressiveRelease={aggressive}  ExcelPID={pid}");

                // なるべく余計なUI要素を作らない（テスト目的）
                app.Visible = false;

                // 何か触ってから正常終了させる（最小）
                var wb = app.Workbooks.Add();
                wb.Close(SaveChanges: false);

                app.Quit();
            }
        }

        // Aggressive=false の場合、ここで GC を促して「すぐ消えるか」を確認
        // （本番設計として GC を呼ぶかどうかは別議論。テストとしては有効）
        if (pid > 0)
        {
            try
            {
                var p = Process.GetProcessById(pid);

                // 少し待ってみる（ここでは最大3秒）
                if (p.WaitForExit(3000))
                    Console.WriteLine($"[TEST] Excel exited. PID={pid}");
                else
                    Console.WriteLine($"[TEST] Excel still running. PID={pid}");
            }
            catch (ArgumentException)
            {
                // 既に終了していると GetProcessById が投げる
                Console.WriteLine($"[TEST] Excel exited (not found). PID={pid}");
            }
        }
    }

}
