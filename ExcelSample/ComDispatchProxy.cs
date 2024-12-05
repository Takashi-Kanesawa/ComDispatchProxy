using System.Reflection;
using System.Runtime.InteropServices;

public class ComDispatchProxy<T> : DispatchProxy where T : class 
{
    private T _comObject;

    public void Initialize(T comObject)
    {
        this._comObject = comObject;
    }

    protected override object? Invoke(MethodInfo? method, object?[]? args)
    {
        return method.Invoke(_comObject, args);
    }

    public static T CreateProxy(T comObject)
    {
        var proxy = Create<T, ComDispatchProxy<T>>() as ComDispatchProxy<T>;
        proxy.Initialize(comObject);

        return proxy as T;
    }

    ~ComDispatchProxy()
    {
        if (_comObject != null && Marshal.IsComObject(_comObject))
        {
#pragma warning disable CA1416 // OS プラットフォームの互換性の警告を無視
            Marshal.ReleaseComObject(_comObject);
#pragma warning restore CA1416
        }
    }

}

