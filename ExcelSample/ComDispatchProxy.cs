using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

public class ComDispatchProxy<T> : DispatchProxy, IDisposable where T : class 
{
    private static readonly Dictionary<ComDispatchProxy<T>, string> _instances = new();

    private Lazy<T> _comObject;
    private Lazy<string> _objectName;
    private Lazy<T> _validProxy;

    public T Proxy { get => this._validProxy.Value; }

    public ComDispatchProxy()
    {
        lock (_instances)
        {
            _instances.Add(this, filterStackTrace(Environment.StackTrace));
        }
#if DEBUG
        AppDomain.CurrentDomain.ProcessExit += (sender, e) =>
        {
            HandleApplicationExit("ProcessExit");
        };

        AppDomain.CurrentDomain.UnhandledException += (sender, e) =>
        {
            Console.WriteLine("[ERROR] Unhandled exception occurred.");
            if (e.ExceptionObject is Exception ex)
            {
                Console.WriteLine($"Exception details: {ex}");
            }
            HandleApplicationExit("UnhandledException");
        };

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

        static void HandleApplicationExit(string reason)
        {
            if (_instances.Count > 0)
            {
                Console.WriteLine($"[WARNING] Application exiting due to {reason}. Some COM objects were not released properly:");
                foreach (var instanceInfo in _instances)
                {
                    Console.WriteLine($" - {instanceInfo.Key._objectName}");
                    Console.WriteLine($"{instanceInfo.Value}");
                    instanceInfo.Key.Dispose();
                }
            }
        }
#endif

        this._comObject = new Lazy<T>(() => throw new InvalidOperationException("COM Object is not initialized."));
        this._objectName = new Lazy<string>(() => throw new InvalidOperationException("Object name is not initialized."));
        this._validProxy = new Lazy<T>(() => throw new InvalidOperationException("COM Object is not initialized."));
    }

    public void Initialize(ComDispatchProxy<T> creatingProxy, T comObject)
    {
        if (creatingProxy == null)
            throw new ArgumentNullException(nameof(creatingProxy), "Creating Proxy cannnot be null.");

        this._comObject = new Lazy<T>(() => comObject ?? throw new ArgumentNullException(nameof(comObject)));
        this._objectName = new Lazy<string>(() => typeof(T).FullName ?? string.Empty);

        if (creatingProxy is not T validProxy)
        {
            throw new ArgumentException($"Creating Proxy is not Proxy of {nameof(T)}.", nameof(creatingProxy));
        }

        this._validProxy = new Lazy<T>(() => validProxy ?? throw new ArgumentNullException(nameof(validProxy)));
    }

    protected override object? Invoke(MethodInfo? method, object?[]? args)
    {
        if (method == null)
            throw new ArgumentNullException(nameof(method));

        if (_comObject.Value == null)
            throw new InvalidOperationException("COM Object is not initialized.");

        Console.WriteLine($"[LOG] Invoking '{method.Name}' on {_objectName.Value} with args: {FormatArgs(args)}");

        try
        {
            // メソッドを実行して戻り値を取得
            var result = method.Invoke(_comObject.Value, args);

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
            return result;
        }
        catch (TargetInvocationException ex)
        {
            Console.WriteLine($"[ERROR] Exception in '{method.Name}': {ex.InnerException?.Message}");
            throw ex.InnerException ?? ex;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[ERROR] Exception in '{method.Name}': {ex.Message}");
            throw;
        }
        
        string FormatArgs(object?[]? args)
        {
            if (args == null || args.Length == 0)
                return "none";
            return string.Join(", ", args.Select(arg => arg?.ToString() ?? "null"));
        }
    }

    private object ProxyFactory(MethodInfo? method, object result)
    {
        if ( method is null || result is null )
        {
            throw new ArgumentNullException($"{nameof(method)} or {nameof(result)} is null.");
        }

        // 型を確認し適切にキャスト
        var typeName = result.GetType().Name;

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

        Console.WriteLine($"[LOG] Can't create proxy for {_objectName.Value}.{method.Name}");
        return result; // 型が異なる場合はそのまま返す

        ComDispatchProxy<TCom> WrapProxy<TCom>(TCom comObject) where TCom : class
        {
            Console.WriteLine($"[LOG] Wrapping returned COM object from '{method.Name}' as '{typeof(TCom)}'.");
            return ComDispatchProxy<TCom>.CreateProxy(comObject) as ComDispatchProxy<TCom>;
        }

    }

    public static ComDispatchProxy<T> CreateProxy(T comObject)
    {
        if (comObject == null)
        {
            throw new ArgumentNullException(nameof(comObject), "Cannot create proxy for a null COM object.");
        }

        var proxy = Create<T, ComDispatchProxy<T>>() as ComDispatchProxy<T>;

        if (proxy == null)
        {
            throw new InvalidOperationException("Failed to create proxy.");
        }

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

        if (_comObject.Value != null && Marshal.IsComObject(_comObject.Value))
        {
#pragma warning disable CA1416 // OS プラットフォームの互換性の警告を無視
            Marshal.ReleaseComObject(_comObject.Value);
#pragma warning restore CA1416
        }

    }

    ~ComDispatchProxy()
    {
        Dispose(false); // ファイナライザから呼び出される場合
    }

}

