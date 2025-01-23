using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

/// <summary>
/// 生のCOMオブジェクトを取得するためのインターフェイス
/// </summary>
public interface IGetRowObject
{
    /// <summary>
    /// 生のCOMオブジェクトを取得します
    /// </summary>
    object RowObject {  get; }
}

/// <summary>
/// COMオブジェクトのプロキシを作成し、メソッド呼び出しを中継するためのクラス。
/// </summary>
/// <typeparam name="T">プロキシ化するCOMオブジェクトの型。</typeparam>
public class ComDispatchProxy<T> : DispatchProxy, IGetRowObject, IDisposable where T : class
{
    // 現在のプロキシインスタンスとそれに関連する情報を保持する辞書
    private static readonly Dictionary<ComDispatchProxy<T>, string> _instances = new();

    // COMオブジェクト、オブジェクト名、有効なプロキシをLazyで保持
    private Lazy<T> _comObject;
    private Lazy<string> _objectName;
    private Lazy<T> _validProxy;

    /// <summary>
    /// プロキシ化されたオブジェクトを取得します。
    /// </summary>
    public T Proxy { get => this._validProxy.Value; }

    public object RowObject => this._comObject.Value;

    /// <summary>
    /// コンストラクタ。プロキシの初期化処理を行います。
    /// </summary>
    public ComDispatchProxy()
    {
        lock (_instances)
        {
            _instances.Add(this, filterStackTrace(Environment.StackTrace));
        }

        // アプリケーション終了時の未解放オブジェクトを警告
        AppDomain.CurrentDomain.ProcessExit += (sender, e) =>
        {
            HandleApplicationExit("ProcessExit");
        };

        // 未処理例外発生時のハンドリング
        AppDomain.CurrentDomain.UnhandledException += (sender, e) =>
        {
            Debug.WriteLine("[ERROR] Unhandled exception occurred.");
            if (e.ExceptionObject is Exception ex)
            {
                Debug.WriteLine($"Exception details: {ex}");
            }
            HandleApplicationExit("UnhandledException");
        };

        // アプリケーション終了時のハンドリング
        static void HandleApplicationExit(string reason)
        {
            if (_instances.Count > 0)
            {
                Debug.WriteLine($"[WARNING] Application exiting due to {reason}. Some COM objects were not released properly:");
                foreach (var instanceInfo in _instances)
                {
                    Debug.WriteLine($" - {instanceInfo.Key._objectName}");
                    Debug.WriteLine($"{instanceInfo.Value}");
                    instanceInfo.Key.Dispose();
                }
            }
        }

        // Lazyのデフォルト初期化
        this._comObject = new Lazy<T>(() => throw new InvalidOperationException("COM Object is not initialized."));
        this._objectName = new Lazy<string>(() => throw new InvalidOperationException("Object name is not initialized."));
        this._validProxy = new Lazy<T>(() => throw new InvalidOperationException("COM Object is not initialized."));
    }


    // スタックトレースのフィルタリングメソッド
    private static string filterStackTrace(string stackTrace)
    {
        Regex[] includesRegex = { new Regex(@"cs:line \d+$") };

        // スタックトレースの行ごとにフィルタリング
        var filteredStackTrace = string.Join(Environment.NewLine, stackTrace
            .Split(new[] { Environment.NewLine }, StringSplitOptions.None)
            .Where(line =>
                includesRegex.Any(regex => regex.IsMatch(line)) &&
                line.StartsWith("   at ComDispatchProxy`") == false));

        return filteredStackTrace;
    }


    /// <summary>
    /// プロキシを初期化します。
    /// </summary>
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

    /// <summary>
    /// メソッド呼び出しを中継します。
    /// </summary>
    protected override object? Invoke(MethodInfo? method, object?[]? args)
    {
        if (method == null)
            throw new ArgumentNullException(nameof(method));

        if (_comObject.Value == null)
            throw new InvalidOperationException("COM Object is not initialized.");

        // 引数リスト内のラッパーを剥がす
        UnwrapPoxiesInArgs(args);

        Debug.WriteLine($"[LOG] Invoking '{method.Name}' on {_objectName.Value} with args: {FormatArgs(args)}");

        try
        {
            // 実際のメソッドを呼び出し、その結果を取得
            var result = method.Invoke(_comObject.Value, args);

            if (result == null)
            {
                Debug.WriteLine($"[LOG] Method '{method.Name}' returned null.");
                return null;
            }

            // COM オブジェクトの場合、適切なプロキシを生成
            if (Marshal.IsComObject(result))
            {
                return ProxyFactory(method, result);
            }

            return result; // 通常の戻り値を返す
        }
        catch (TargetInvocationException ex)
        {
            Debug.WriteLine($"[ERROR] Exception in '{method.Name}': {ex.InnerException?.Message}");
            throw ex.InnerException ?? ex;
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[ERROR] Exception in '{method.Name}': {ex.Message}");
            throw;
        }

        // ローカル関数：引数のフォーマット処理
        string FormatArgs(object?[]? args)
        {
            if (args == null || args.Length == 0)
                return "none";
            return string.Join(", ", args.Select(arg => arg?.ToString() ?? "null"));
        }

        // ローカル関数：引数のCOMオブジェクトをDispatchProxyから生のCOMオブジェクトに置換
        void UnwrapPoxiesInArgs(object?[]? args)
        {
            if (args == null)
                return;

            for (int i = 0; i < args.Length; i++)
            {
                if (args[i] is not IGetRowObject proxy)
                    continue;

                // IGetRowObject を実装している場合は RowObject と入れ替える
                args[i] = proxy.RowObject;
            }
        }
    }

    /// <summary>
    /// COMオブジェクトに基づいて適切なプロキシを生成します。
    /// </summary>
    private object ProxyFactory(MethodInfo? method, object result)
    {
        if (method is null || result is null)
        {
            throw new ArgumentNullException($"{nameof(method)} or {nameof(result)} is null.");
        }

        Debug.WriteLine($"result.GetType() = {result.GetType()}");

        // 型ごとにプロキシを生成
        switch (result)
        {
            // 基本操作で頻繁に使用されるオブジェクト
            case Excel.Range range: return WrapProxy(range);    // Rangeは最も使用頻度が高そう
            case Excel.Application application: return WrapProxy(application);
            case Excel.Workbooks workbooks: return WrapProxy(workbooks);
            case Excel.Workbook workbook: return WrapProxy(workbook);
            case Excel.Sheets sheets: return WrapProxy(sheets);
            case Excel.Worksheet worksheet: return WrapProxy(worksheet);

            // スタイル・グラフ関連の末端オブジェクト
            case Excel.Font font: return WrapProxy(font);
            case Excel.Border border: return WrapProxy(border);
            case Excel.Borders borders: return WrapProxy(borders);
            case Excel.Chart chart: return WrapProxy(chart);
            case Excel.ChartObject chrObj: return WrapProxy(chrObj);
            case Excel.Axis axis: return WrapProxy(axis);
            case Excel.Point point: return WrapProxy(point);
            case Excel.Series series: return WrapProxy(series);

            // グラフ関連のコレクション
            case Excel.ChartObjects chrObjs: return WrapProxy(chrObjs);
            case Excel.SeriesCollection seriesCollction: return WrapProxy(seriesCollction);
            case Excel.Axes axes: return WrapProxy(axes);
            case Excel.Points points: return WrapProxy(points);

            // 使用頻度の少なそうな他の要素
            case Excel.Shape shape: return WrapProxy(shape);
            case Excel.Shapes shapes: return WrapProxy(shapes);
            case Excel.ListObject lstObj: return WrapProxy(lstObj);
            case Excel.ListObjects lstObjs: return WrapProxy(lstObjs);
            case Excel.Hyperlink hLink: return WrapProxy(hLink);
            case Excel.Hyperlinks hLinks: return WrapProxy(hLinks);
            case Excel.Name name: return WrapProxy(name);
            case Excel.Names names: return WrapProxy(names);
            case Excel.PivotTable pt: return WrapProxy(pt);
            case Excel.PivotTables pts: return WrapProxy(pts);
            case Excel.Interior interior: return WrapProxy(interior);
            case Excel.Pictures pictures: return WrapProxy(pictures);

            // その他
            case Excel.Windows windows: return WrapProxy(windows);
            case Excel.Window window: return WrapProxy(window);
            case Excel.AutoFilter af: return WrapProxy(af);
            case Excel.Filters filters: return WrapProxy(filters);
            case Excel.PageSetup pageSetup: return WrapProxy(pageSetup);
            case Excel.QueryTable queryTable: return WrapProxy(queryTable);
            case Excel.TableStyle tableStyle: return WrapProxy(tableStyle);
            case Excel.PivotCache pivotCache: return WrapProxy(pivotCache);
            case Excel.PivotField pivotField: return WrapProxy(pivotField);
            case Excel.PivotItem pivotItem: return WrapProxy(pivotItem);
            case Excel.Trendlines trendlines: return WrapProxy(trendlines);
            case Excel.DataLabels dataLabels: return WrapProxy(dataLabels);
            case Excel.DataLabel dataLabel: return WrapProxy(dataLabel);
            case Excel.Legend legend: return WrapProxy(legend);
            case Excel.Style style: return WrapProxy(style);
            case Excel.Validation validation: return WrapProxy(validation);
            case Excel.FormatCondition formatCondition: return WrapProxy(formatCondition);
            case Excel.FormatConditions formatConditions: return WrapProxy(formatConditions);
            case Excel.Comment comment: return WrapProxy(comment);
            case Excel.Comments comments: return WrapProxy(comments);
            case Excel.GroupBoxes groupBoxes: return WrapProxy(groupBoxes);
        }
        
        Debug.WriteLine($"[LOG] Can't create proxy for {_objectName.Value}.{method.Name}");
        return result; // 型が一致しない場合そのまま返す

        // ローカル関数：COMオブジェクトをDispatchProxyでラップする。
        ComDispatchProxy<TCom> WrapProxy<TCom>(TCom comObject) where TCom : class
        {
            Debug.WriteLine($"[LOG] Wrapping returned COM object from '{method.Name}' as '{typeof(TCom)}'.");
            return ComDispatchProxy<TCom>.CreateProxy(comObject) as ComDispatchProxy<TCom>;
        }
    }

    /// <summary>
    /// 新しいプロキシを生成します。
    /// </summary>
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
        Debug.WriteLine($"[LOG] ComDispatchProxy is created as '{typeof(T)}' at \n{filterStackTrace(Environment.StackTrace)}");

        return proxy;
    }

    /// <summary>
    /// プロキシを解放し、関連リソースを破棄します。
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this); // デストラクタをスキップ
    }

    protected virtual void Dispose(bool disposing)
    {
        Debug.WriteLine($"{this._objectName} is disposing.");

        lock (_instances)
        {
            _instances.Remove(this);
        }

        if (_comObject.Value != null && Marshal.IsComObject(_comObject.Value))
        {
#pragma warning disable CA1416 // OS互換性警告を無視
            Marshal.ReleaseComObject(_comObject.Value);
#pragma warning restore CA1416
        }
    }

    /// <summary>
    /// デストラクタ。Disposeを呼び出してリソースを解放します。
    /// </summary>
    ~ComDispatchProxy()
    {
        Dispose(false); // ファイナライザから呼び出し
    }
}
