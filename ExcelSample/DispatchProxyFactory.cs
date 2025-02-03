using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;

public static class DispatchProxyFactory
{
    private static readonly Dictionary<Type, Func<object, object>> ProxyCreators = new();

    /// <summary>
    /// XML ファイルから対象の COM インターフェイスを読み込み、ProxyCreators を初期化
    /// </summary>
    public static void InitializeFromXml(string xmlPath)
    {
        lock (ProxyCreators)
        {
            ProxyCreators.Clear();

            var xml = XDocument.Load(xmlPath);

            // 生成対象のインターフェイスリストを取得
            var generateInterfaces = xml
                .Descendants("Generate")
                .Descendants("Interface")
                .Select(x => x.Value)
                .ToHashSet(); // 重複防止

            var allTypes = typeof(Excel.Application).Assembly.GetTypes()
                .Where(t => t.IsInterface && t.Namespace == "Microsoft.Office.Interop.Excel");

            foreach (var type in allTypes)
            {
                if (type.FullName == null) continue;
                                                                                                                                                                                                                                                                                                   
                if (generateInterfaces.Contains(type.FullName))
                {
                    ProxyCreators[type] = obj => CreateProxyForType(type, obj);
                    Debug.WriteLine($"[LOG] Registered proxy for {type.FullName}");
                }
            }
        }                                                                                                                                                                                                                                                                                                                                                                                                                                
    }
    /*
        { typeof(Excel.Application), obj => ComDispatchProxy<Excel.Application>.CreateProxy((Excel.Application)obj, "Excel.Application") },
        { typeof(Excel.Workbooks), obj => ComDispatchProxy<Excel.Workbooks>.CreateProxy((Excel.Workbooks)obj, "Excel.Workbooks") },
     */


    /// <summary>                                                                                                                                                                                                                                                                                                                                                                                                                            
    /// COM オブジェクトを判定し、適切なプロキシを生成します。
    /// </summary>
    public static object? CreateProxy(object comObject)
    {
        if (comObject == null) throw new ArgumentNullException(nameof(comObject));

        foreach (var (type, creator) in ProxyCreators)
        {
            if (IsComObjectOfType(comObject, type))
            {
                Debug.WriteLine($"[LOG] Creating proxy for type: {type.FullName}");
                return creator(comObject);
            }
        }

        Debug.WriteLine($"[LOG] No matching type found for COM object. Returning original object.");
        return comObject; // 一致する型がない場合そのまま返す
    }

    /// <summary>
    /// COM オブジェクトが特定の型を実装しているか確認します（QueryInterface を使用）。
    /// </summary>
    private static bool IsComObjectOfType(object comObject, Type interfaceType)
    {
        IntPtr comPointer = IntPtr.Zero;

        try
        {
            comPointer = Marshal.GetIUnknownForObject(comObject);
            Guid interfaceGuid = interfaceType.GUID;
            IntPtr queriedPointer;

            int hr = Marshal.QueryInterface(comPointer, ref interfaceGuid, out queriedPointer);
            if (hr == 0) // S_OK: 型が一致
            {
                Marshal.Release(queriedPointer);
                return true;
            }
        }
        finally
        {
            if (comPointer != IntPtr.Zero)
            {
                Marshal.Release(comPointer);
            }
        }

        return false;
    }

    /// <summary>
    /// 指定された型のプロキシを作成
    /// </summary>
    private static object CreateProxyForType(Type interfaceType, object comObject)
    {
        var proxyType = typeof(ComDispatchProxy<>).MakeGenericType(interfaceType);
        var createMethod = proxyType.GetMethod("CreateProxy",
            BindingFlags.Static | BindingFlags.Public,
            null,
            new Type[] { interfaceType }, 
            null);

        if (createMethod == null)
        {
            throw new InvalidOperationException($"Failed to locate 'CreateProxy' method on {proxyType.FullName}");
        }

        var methodInfo = createMethod.Invoke(null, new object[] { comObject });

        if ( methodInfo == null)
        {
            throw new InvalidOperationException($"Unhandled COM object type: {proxyType.FullName}");
        }
        return methodInfo;
    }
}
