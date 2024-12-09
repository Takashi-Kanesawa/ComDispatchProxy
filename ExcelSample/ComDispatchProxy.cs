using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

public class ComDispatchProxy<T> : DispatchProxy, IDisposable where T : class 
{
    private static readonly Dictionary<ComDispatchProxy<T>, string> _instances = new();

    private T _comObject;
    private string _objectName;
    private T _validProxy;

    public T Proxy => this._validProxy;

    public ComDispatchProxy()
    {
        lock (_instances)
        {
            _instances.Add(this, filterStackTrace(Environment.StackTrace));

	        string filterStackTrace(string stackTrace)
	        {
	            Regex[] includesRegex = { new Regex(@"cs:line \d+$") };
	
	            // StackTraceを行ごとに分割し、フィルタリング
	            var filteredStackTrace = string.Join(Environment.NewLine, stackTrace
	                .Split(new[] { Environment.NewLine }, StringSplitOptions.None)
	                .Where(line =>
	                    includesRegex.Any(regex => regex.IsMatch(line)) &&
	                    line.StartsWith("   at ComDispatchProxy`") == false));
	
	            return filteredStackTrace;
	        }
        }

        AppDomain.CurrentDomain.ProcessExit += (sender, e) =>
        {
            if (_instances.Count > 0)
            {
                Console.WriteLine($"[WARNING] Application exiting. Some COM objects were not released properly:");
                foreach (var instanceInfo in _instances)
                {
                    Console.WriteLine($" - {instanceInfo.Key._objectName}");
                    Console.WriteLine($"{instanceInfo.Value}");
                    instanceInfo.Key.Dispose();
                }
            }
        };

    }

    public void Initialize(ComDispatchProxy<T> creatingProxy, T comObject)
    {
        this._comObject = comObject;
        this._objectName = typeof(T).FullName;

        this._validProxy = creatingProxy as T;
    }

    protected override object? Invoke(MethodInfo? method, object?[]? args)
    {
        Console.WriteLine($"[LOG] Invoking '{method.Name}' on {_objectName} with args: {FormatArgs(args)}");

        // メソッドを実行して戻り値を取得
        var result = method.Invoke(_comObject, args);

        // 戻り値が null の場合
        if (result == null)
        {
            Console.WriteLine($"[LOG] Method '{method.Name}' returned null.");
            return null;
        }

        // COM オブジェクトの場合、型情報を使ってプロキシを生成
        if (Marshal.IsComObject(result))
        {
            return ProxyFactory(method, result);
        }

        // 戻り値が COM オブジェクトでない場合
        Console.WriteLine($"[LOG] Method  '{method.Name}' returned '{result.GetType().Name}'.");
        return result;

        string FormatArgs(object?[]? args)
        {
            if (args == null || args.Length == 0)
                return "none";
            return string.Join(", ", args.Select(arg => arg?.ToString() ?? "null"));
        }
    }

    private object ProxyFactory(MethodInfo? method, object result)
    {
        switch (result)
        {
            case Excel.Application app: return WrapProxy(app);
            case Excel.Workbooks wbs: return WrapProxy(wbs);
            case Excel.Workbook wb: return WrapProxy(wb);
            case Excel.Sheets wss: return WrapProxy(wss);
            case Excel.Worksheet ws: return WrapProxy(ws);
            case Excel.Range range: return WrapProxy(range);
            case Excel.Windows wins: return WrapProxy(wins);
            case Excel.Window win: return WrapProxy(win);
        }

        Console.WriteLine($"[LOG] Can't create proxy for {_objectName}.{method.Name}");
        return result; // 型が異なる場合はそのまま返す

        ComDispatchProxy<TCom> WrapProxy<TCom>(TCom comObject) where TCom : class
        {
            Console.WriteLine($"[LOG] Wrapping returned COM object from '{method.Name}' as '{typeof(TCom)}'.");
            return ComDispatchProxy<TCom>.CreateProxy(comObject) as ComDispatchProxy<TCom>;
        }
    }

    public static ComDispatchProxy<T> CreateProxy(T comObject)
    {
        var proxy = Create<T, ComDispatchProxy<T>>() as ComDispatchProxy<T>;
        proxy.Initialize(proxy, comObject);

        return proxy;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this); // デストラクタをスキップする
    }

    protected virtual void Dispose(bool disposing)
    {
        lock (_instances)
        {
            _instances.Remove(this);
        }

        if (_comObject != null && Marshal.IsComObject(_comObject))
        {
#pragma warning disable CA1416 // OS プラットフォームの互換性の警告を無視
            Marshal.ReleaseComObject(_comObject);
#pragma warning restore CA1416
        }

    }

    ~ComDispatchProxy()
    {
        Dispose(false); // ファイナライザから呼び出される場合
    }

}

