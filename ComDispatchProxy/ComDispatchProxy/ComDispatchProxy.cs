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
    private T _rcw = null!;
    private string _objectName = null!;
    private T _validProxy = null!;
    private IComProxyFactory _proxyFactory = null!;
    private readonly Dictionary<IComDispatchProxy, string> _childProxies = new();
    private ComProxySession _session = null!;
    ComProxySession IComProxySessionProvider.Session => _session;
    #endregion

    #region IComDispatchProxyの実装
    private bool _isDisposed = false;
    public bool IsDisposed => this._isDisposed;

    private bool _isRoot = false;
    public bool IsRoot => this._isRoot;

    public T Proxy
    {
        get
        {
            if (this.IsDisposed)
            {
                throw new InvalidOperationException($"Proxy has been released : {this._objectName}");
            }

            return this._validProxy;
        }
    }
    public IComProxyFactory ProxyFactory => this._proxyFactory;

    object? IComDispatchProxy.RawRcw  => this.IsDisposed ? null : this._rcw;

    public void AddChild(IComDispatchProxy childObject)
    {
        lock (_childProxies)
        {
            if (_childProxies.ContainsKey(childObject) == false)
            {
                _childProxies.Add(childObject, filterStackTrace(Environment.StackTrace));
            }
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
    }

    /// <summary>
    /// プロキシを初期化します。
    /// </summary>
    internal void Initialize(
        IComProxyFactory proxyFactory,
        T validProxy,
        T rcw,
        ComProxySession session,
        bool isRoot)
    {
        _proxyFactory = proxyFactory ?? throw new ArgumentNullException(nameof(proxyFactory));
        _rcw = rcw ?? throw new ArgumentNullException(nameof(rcw));
        _objectName = typeof(T).FullName ?? string.Empty;

        _validProxy = validProxy ?? throw new ArgumentNullException(nameof(validProxy));
        _session = session ?? throw new ArgumentNullException(nameof(session));
        _isRoot = isRoot;
    }

    public static ComDispatchProxy<T> CreateProxy(ComProxyFactoryBase proxyFactory, T comObject)
    {
        return CreateProxy(proxyFactory, comObject, aggressiveReleaseComObjects: true);
    }

    public static ComDispatchProxy<T> CreateProxy(ComProxyFactoryBase proxyFactory, T comObject, bool aggressiveReleaseComObjects)
    {
        return CreateProxyInternal(
            proxyFactory,
            comObject,
            parentObject: null,
            session: new ComProxySession(aggressiveReleaseComObjects));
    }

    // 既存：壊さない（Factory側のreflection用）
    public static ComDispatchProxy<T> CreateProxy(ComProxyFactoryBase proxyFactory, T comObject, IComDispatchProxy? parentObject)
    {
        if (parentObject is null)
        {
            throw new ArgumentException(
                "Root proxy must be created by CreateProxy(factory, comObject) or CreateProxy(factory, comObject, aggressiveReleaseComObjects).",
                nameof(parentObject));
        }

        return CreateProxyInternal(proxyFactory, comObject, parentObject, parentObject.Session, isRoot : false);
    }

    private static ComDispatchProxy<T> CreateProxyInternal(
        ComProxyFactoryBase proxyFactory,
        T comObject,
        IComDispatchProxy? parentObject,
        ComProxySession session,
        bool isRoot = true)
    {
        if (proxyFactory is null) throw new ArgumentNullException(nameof(proxyFactory));
        if (comObject is null) throw new ArgumentNullException(nameof(comObject));

        // DispatchProxy を T として生成
        T proxyAsT = DispatchProxy.Create<T, ComDispatchProxy<T>>();

        // 実装インスタンス（DispatchProxy本体）を取り出す
        var impl = (ComDispatchProxy<T>)(object)proxyAsT;

        // Initialize は代入しかしない
        impl.Initialize(proxyFactory, proxyAsT, comObject, session, isRoot);

        if (isRoot)
        {
            ComProxyAppDomainHook.RegisterRoot(impl);
        }

        ComProxyLog.Write($"[ComDispatchProxy CREATED] ComDispatchProxy is created as '{typeof(T)}'");
        return impl;
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

        if (_rcw == null)
            throw new InvalidOperationException("COM Object is not initialized.");

        // 引数内のプロキシを解除し、元の COM オブジェクトに戻す
        UnwrapProxiesInArgs(args);

        ComProxyLog.Write($"[ComDispatchProxy LOG] Invoking '{method.Name}' on {_objectName} with args: {FormatArgs(args)}");

        try
        {
            var result = method.Invoke(_rcw, args);
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
                    args[i] = proxy.RawRcw;
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
            if (this.IsRoot && this.IsDisposed == false)
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

        if (this._session.AggressiveReleaseComObjects)
        {
            GC.SuppressFinalize(this); // デストラクタをスキップ
        }
    }

    protected void Dispose(bool disposing)
    {
        lock (this._childProxies)
        {
            // もう Release 済みなら何もしない（多重呼び出しガード）
            if (this.IsDisposed)
            {
                return;
            }

            try
            {
                if (disposing)
                {
                    // マネージ側の子プロキシだけ先に片付ける
                    foreach (var child in _childProxies.Keys.ToList())
                    {
                        child.Dispose();
                    }
                    this._childProxies.Clear();

                    if (this.IsRoot)
                    {
                        this._session.Dispose();
                    }
                }

                if (Marshal.IsComObject(this._rcw))
                {
                    if (this._session.AggressiveReleaseComObjects)
                    {
#pragma warning disable CA1416 // OS互換性警告を無視
                        Marshal.ReleaseComObject(this._rcw);
#pragma warning restore CA1416
                        ComProxyLog.Write($"[ComDispatchProxy RELEASED]{this._objectName} has been released.");
                    }
                    else
                    {
                        ComProxyLog.Write($"[ComDispatchProxy DISPOSED]{this._objectName} disposed without ReleaseComObject (AggressiveReleaseComObjects=false).");
                    }
                }
            }
            finally
            {
                this._isDisposed = true;
                // 参照も切ってGC対象に寄せる（Aggressive/非Aggressiveどちらでも有益）
                this._rcw = default!;
                this._proxyFactory = default!;
                this._validProxy = default!;
            }
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