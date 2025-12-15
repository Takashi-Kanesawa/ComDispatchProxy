#region usings
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

#endregion

namespace ComDispatchProxy;

#region ComDispatchProxy クラス定義
/// <summary>
/// COMオブジェクトのプロキシを作成し、メソッド呼び出しを中継するためのクラス。
/// </summary>
/// <typeparam name="T">プロキシ化するCOMオブジェクトの型。</typeparam>
public class ComDispatchProxy<T> : DispatchProxy, IComDispatchProxy where T : class
{
    #region private field
    private T _comObject = null!;
    private string _objectName = null!;
    private T _validProxy = null!;
    private IComProxyFactory _proxyFactory = null!;
    private readonly Dictionary<IComDispatchProxy, string> _childProxies = new();
    #endregion

    #region IComDispatchProxyの実装
    private bool _wasReleased = false;
    public bool WasReleased => this._wasReleased;

    private bool _disposed = false;
    public bool DIsposed => this._disposed;

    private bool _isRoot = false;
    public bool IsRoot => this._isRoot;

    public T Proxy
    {
        get
        {
            if (this.WasReleased)
            {
                throw new InvalidOperationException($"Proxy has been released : {this._objectName}");
            }

            return this._validProxy;
        }
    }
    public IComProxyFactory ProxyFactory => this._proxyFactory;

    public object? RowObject => this.WasReleased ? null : this._comObject;

    public void AddChild(IComDispatchProxy childObject)
    {
        lock (_childProxies)
        {
            _childProxies.Add(childObject, filterStackTrace(Environment.StackTrace));
        }
    }

    public void RemoveChild(IComDispatchProxy childObject)
    {
        lock (_childProxies)
        {
            _childProxies.Remove(childObject);
        }
    }
    #endregion

    #region コンストラクタ及び初期化処理
    /// <summary>
    /// コンストラクタ。プロキシの初期化処理を行います。
    /// </summary>
    public ComDispatchProxy()
    {
        // アプリケーション終了時の未解放オブジェクトを警告
        AppDomain.CurrentDomain.ProcessExit += (sender, e) =>
        {
            HandleApplicationExit("ProcessExit", this);
        };

        // 未処理例外発生時のハンドリング
        AppDomain.CurrentDomain.UnhandledException += (sender, e) =>
        {
            ComProxyLog.Write("[ComDispatchProxy ERR] Unhandled exception occurred.");
            if (e.ExceptionObject is Exception ex)
            {
                ComProxyLog.Write($"Exception details: {ex}");
            }
            HandleApplicationExit("UnhandledException", this);
        };

        // アプリケーション終了時のハンドリング
        static void HandleApplicationExit(string reason, ComDispatchProxy<T> proxy)
        {
            proxy.DisposeIfRoot();
        }
    }

    /// <summary>
    /// プロキシを初期化します。
    /// </summary>
    internal void Initialize(IComProxyFactory proxyFactory, object? parentProxy, ComDispatchProxy<T> creatingProxy, T comObject)
    {
        if (creatingProxy == null)
        {
            throw new ArgumentNullException(nameof(creatingProxy), "Creating Proxy cannnot be null.");
        }

        this._proxyFactory = proxyFactory ?? throw new ArgumentNullException(nameof(proxyFactory));
        this._comObject = comObject ?? throw new ArgumentNullException(nameof(comObject));
        this._objectName = typeof(T).FullName ?? string.Empty;

        if (creatingProxy is not T validProxy)
        {
            throw new ArgumentException($"Creating Proxy is not Proxy of {nameof(T)}.", nameof(creatingProxy));
        }

        this._validProxy = validProxy;

        // 最上位（Excel.Application）の親は無いので
        if (parentProxy is null)
        {
            this._isRoot = true;

            return;
        }

        if (parentProxy is not IComDispatchProxy validParentProxy)
        {
            throw new ArgumentException($"Parent object must be IComDispatchProxy when not null.", nameof(parentProxy));
        }
    }

    public static ComDispatchProxy<T> CreateProxy(IComProxyFactory proxyFactory, T comObject, object? parentObject = null)
    {
        if( proxyFactory == null)
        {
            throw new ArgumentNullException(nameof(proxyFactory), "A proxyFactory is required to create a proxy.");
        }

        if (comObject == null)
        {
            throw new ArgumentNullException(nameof(comObject), "Cannot create proxy for a null COM object.");
        }

        var proxy = Create<T, ComDispatchProxy<T>>() as ComDispatchProxy<T>;
                                                                                                                                                               
        if (proxy == null)
        {
            throw new InvalidOperationException("Failed to create proxy.");
        }

        proxy.Initialize(proxyFactory, parentObject, proxy, comObject);

        ComProxyLog.Write($"[ComDispatchProxy CREATED] ComDispatchProxy is created as '{typeof(T)}'");
        return proxy;
    }
    #endregion

    #region COMインターフェイスの呼び出しの中継と、Factory

    /// <summary>
    /// メソッド呼び出しを中継し、取得した COM オブジェクトを適切にラップする。
    /// </summary>
    protected override object? Invoke(MethodInfo? method, object?[]? args)
    {
        if (method == null)
            throw new ArgumentNullException(nameof(method));

        if (_comObject == null)
            throw new InvalidOperationException("COM Object is not initialized.");

        // 引数内のプロキシを解除し、元の COM オブジェクトに戻す
        UnwrapProxiesInArgs(args);

        ComProxyLog.Write($"[ComDispatchProxy LOG] Invoking '{method.Name}' on {_objectName} with args: {FormatArgs(args)}");

        try
        {
            var result = method.Invoke(_comObject, args);
            if (result == null)
            {
                ComProxyLog.Write($"[ComDispatchProxy LOG] Method '{method.Name}' returned null.");
                return null;
            }

            return WrapProxyIfComObject(result);
        }
        catch (TargetInvocationException ex)
        {
            ComProxyLog.Write($"[ComDispatchProxy ERR] Exception in '{method.Name}': {ex.InnerException?.Message}");
            throw ex.InnerException ?? ex;
        }
        catch (Exception ex)
        {
            ComProxyLog.Write($"[ComDispatchProxy ERR] Exception in '{method.Name}': {ex.Message}");
            throw;
        }

        // -----------------------------------------------------------
        // ローカル関数：引数のフォーマット処理（デバッグログ用）
        string FormatArgs(object?[]? args)
        {
            if (args == null || args.Length == 0)
                return "none";

            return string.Join(", ", args.Select(arg => arg?.ToString() ?? "null"));
        }

        // -----------------------------------------------------------
        // ローカル関数：引数の COM オブジェクトを生のオブジェクトに戻す
        void UnwrapProxiesInArgs(object?[]? args)
        {
            if (args == null) return;

            for (int i = 0; i < args.Length; i++)
            {
                if (args[i] is IComDispatchProxy proxy)
                {
                    args[i] = proxy.RowObject;
                }
            }
        }

        // -----------------------------------------------------------
        // ローカル関数：メソッドの返値がCOMオブジェクトの場合はProxy化する
        object? WrapProxyIfComObject(object result)
        {
            // 返り値が COM オブジェクトの場合、プロキシを作成
            if (Marshal.IsComObject(result))
            {
                ComProxyLog.Write($"[ComDispatchProxy LOG] Wrapping returned COM object from '{method.Name}'.");

                var childProxy = this.ProxyFactory.CreateProxyByFactoryFunction(result, this);
                if (childProxy is IComDispatchProxy dispatchProxy)
                {
                    this.AddChild(dispatchProxy);
                }
                return childProxy;
            }

            return result; // 通常のオブジェクトならそのまま返す
        }
    }

    #endregion

    #region ヘルパ

    private void DisposeIfRoot()
    {
        lock (this._childProxies)
        {
            if (this.IsRoot && this.WasReleased == false)
            {
                this.Dispose();
            }
        }
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
    #endregion

    #region IDisposable
    /// <summary>
    /// プロキシを解放し、関連リソースを破棄します。
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this); // デストラクタをスキップ
    }

    protected void Dispose(bool disposing)
    {
        lock (_childProxies)
        {
            // もう Release 済みなら何もしない（多重呼び出しガード）
            if (this.WasReleased)
            {
                return;
            }

            if (disposing)
            {
                // マネージ側の子プロキシだけ先に片付ける
                foreach (var child in _childProxies.Keys.ToList())
                {
                    child.Dispose();
                }
                _childProxies.Clear();
            }

            // COM じゃなければ何もしない
            if (!Marshal.IsComObject(_comObject))
            {
                _wasReleased = true; // 一応フラグだけ立てるならここ
                return;
            }

#pragma warning disable CA1416 // OS互換性警告を無視
            Marshal.ReleaseComObject(_comObject);
#pragma warning restore CA1416

            _wasReleased = true;
            ComProxyLog.Write($"[ComDispatchProxy RELEASED]{_objectName} has been released.");
        }
    }

    /// <summary>
    /// デストラクタ。Disposeを呼び出してリソースを解放します。
    /// </summary>
    ~ComDispatchProxy()
    {
        Dispose(false); // ファイナライザから呼び出し
    }
    #endregion
}
#endregion