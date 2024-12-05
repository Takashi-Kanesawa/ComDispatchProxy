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
        // メソッドを実行して戻り値を取得
        var result = method.Invoke(_comObject, args);

        // 戻り値が null の場合
        if (result == null)
        {
            return null;
        }

        // COM オブジェクトの場合、型情報を使ってプロキシを生成
        if (Marshal.IsComObject(result))
        {
            return WrapComObjectWithProxy(method, result);
        }

        // 戻り値が COM オブジェクトでない場合
        return result;
    }

    private object WrapComObjectWithProxy(MethodInfo? method, object? result)
    {
        switch (result)
        {
            case Excel.Application app: return WrapProxy(app);
            case Excel.Workbooks wbs: return WrapProxy(wbs);
            case Excel.Workbook wb: return WrapProxy(wb);
            case Excel.Sheets wss: return WrapProxy(wss);
            case Excel.Worksheet ws: return WrapProxy(ws);
            case Excel.Range range: return WrapProxy(range);
        }

        return result; // 型が異なる場合はそのまま返す

        static ComDispatchProxy<TCom> WrapProxy<TCom>(TCom comObject) where TCom : class
        {
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

