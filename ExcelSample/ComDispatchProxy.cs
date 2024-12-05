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
}

