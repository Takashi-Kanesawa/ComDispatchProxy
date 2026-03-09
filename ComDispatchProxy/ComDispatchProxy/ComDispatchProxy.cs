#region usings
using System.Reflection;
using System.Runtime.InteropServices;

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
    private readonly object _disposeGate = new(); // ChildProxies廃止後のDisposeガード用
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
            if (this.IsDisposed || this._session.IsDisposed)
            {
                throw new ObjectDisposedException(this._objectName);
            }

            return this._validProxy;
        }
    }
    public IComProxyFactory ProxyFactory => this._proxyFactory;

    object? IComDispatchProxy.RawRcw
    {
        get
        {
            if (this.IsDisposed || this._session.IsDisposed)
            {
                throw new ObjectDisposedException(this._objectName);
            }

            return this._rcw;
        }
    }

    string IComDispatchProxy.ComObjectName => this._objectName;
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

    public static ComDispatchProxy<T> CreateProxy(IComProxyFactory proxyFactory, T comObject)
    {
        return CreateProxy(proxyFactory, comObject, aggressiveReleaseComObjects: true);
    }

    public static ComDispatchProxy<T> CreateProxy(
        IComProxyFactory proxyFactory, 
        T comObject, 
        bool aggressiveReleaseComObjects)

    {
        if (proxyFactory is null)
        {
            throw new ArgumentNullException(nameof(proxyFactory));
        }

        if (comObject is null)
        {
            throw new ArgumentNullException(nameof(comObject));
        }

        if (Marshal.IsComObject(comObject) == false)
        {
            throw new ArgumentException("comObject must be a COM RCW.", nameof(comObject));
        }

        return CreateProxyInternal(
            proxyFactory,
            comObject,
            parentObject: null,
            session: new ComProxySession(aggressiveReleaseComObjects));
    }

    // 既存：壊さない（Factory側のreflection用）
    public static ComDispatchProxy<T> CreateProxy(
        IComProxyFactory proxyFactory, 
        T comObject, 
        IComDispatchProxy parentObject)
    {
        if (parentObject is null)
        {
            throw new ArgumentException(
                "Child proxy must be created by CreateProxy(factory, comObject, parentObject). " +
                "For root, use CreateProxy(factory, comObject) or CreateProxy(factory, comObject, aggressiveReleaseComObjects).",
                nameof(parentObject));
        }

        return CreateProxyInternal(proxyFactory, comObject, parentObject, parentObject.Session, isRoot : false);
    }

    private static ComDispatchProxy<T> CreateProxyInternal(
        IComProxyFactory proxyFactory,
        T comObject,
        IComDispatchProxy? parentObject,
        ComProxySession session,
        bool isRoot = true)
    {
        if (proxyFactory is null) 
        {
            throw new ArgumentNullException(nameof(proxyFactory)); 
        }
        
        if (comObject is null)
        {
            throw new ArgumentNullException(nameof(comObject));
        }

        // DispatchProxy を T として生成
        T proxyAsT = DispatchProxy.Create<T, ComDispatchProxy<T>>();

        // 実装インスタンス（DispatchProxy本体）を取り出す
        var impl = (ComDispatchProxy<T>)(object)proxyAsT;

        // Initialize は代入しかしない
        impl.Initialize(proxyFactory, proxyAsT, comObject, session, isRoot);

        if (isRoot && Marshal.IsComObject(comObject))
        {
            session.TrackOrReleaseDuplicate(comObject, $"ROOT:{typeof(T).FullName}");
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
    protected override object? Invoke(MethodInfo? methodInfo, object?[]? args)
    {
        if (methodInfo == null)
        {
            throw new ArgumentNullException(nameof(methodInfo));
        }

        if (this.IsDisposed || this._session.IsDisposed)
        {
            throw new ObjectDisposedException(this._objectName);
        }

        if (_rcw == null)
        {
            throw new InvalidOperationException("COM Object is not initialized.");
        }

        // 引数内のプロキシを解除し、元の COM オブジェクトに戻す
        UnwrapProxiesInArgs(args);

        ComProxyLog.Write($"Invoking '{methodInfo.Name}' on {_objectName} with args: {FormatArgs(args)}");

        try
        {
            var result = methodInfo.Invoke(_rcw, args);
            if (result == null)
            {
                ComProxyLog.Write($"Method '{methodInfo.Name}' returned null.");
                return null;
            }

            // foreach 対策：列挙子(IEnumerator)は Current が Invoke を通らないことがあるためラップする
            if (result is System.Collections.IEnumerator ie)
            {
                return new ComEnumeratorWrapper(
                    ie,
                    parent: (IComDispatchProxy)this,
                    ctxBase: $"{_objectName}.{methodInfo.Name}");
            }

            // 方針：戻り値COMのみ Track（out/ref・COM配列は無視）
            if (Marshal.IsComObject(result))
            {
                // ここで proxy化できる/できないに関係なく、RCWをセッション管理下へ
                this._session.TrackOrReleaseDuplicate(result, $"{_objectName}.{methodInfo.Name}");
            }

            return WrapProxyIfComObject(result);
        }
        catch (TargetInvocationException ex)
        {
            ComProxyLog.Write($"[ComDispatchProxy ERR] Exception in '{methodInfo.Name}': {ex.InnerException?.Message}");
            throw ex.InnerException ?? ex;
        }
        catch (Exception ex)
        {
            ComProxyLog.Write($"[ComDispatchProxy ERR] Exception in '{methodInfo.Name}': {ex.Message}");
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
            if (args == null)
            {
                return;
            }

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
                ComProxyLog.Write($"Wrapping returned COM object from '{methodInfo.Name}'.");

                var childProxy = this.ProxyFactory.CreateProxyByFactoryFunction(result, this);
                return childProxy;
            }

            return result; // 通常のオブジェクトならそのまま返す
        }
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
        lock (this._disposeGate)
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
                    if (this.IsRoot && this._session.IsDisposed == false)
                    {
                        this._session.Dispose();
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