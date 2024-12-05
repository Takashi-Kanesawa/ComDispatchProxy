using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

public class ComDispatchProxy<T> : DispatchProxy, IDisposable where T : class 
{
    private T _comObject;
    private T _validProxy;

    public T Proxy => this._validProxy;

    public void Initialize(ComDispatchProxy<T> creatingProxy, T comObject)
    {
        this._comObject = comObject;

        this._validProxy = creatingProxy as T;
    }

    protected override object? Invoke(MethodInfo? method, object?[]? args)
    {
        return method.Invoke(_comObject, args);
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

