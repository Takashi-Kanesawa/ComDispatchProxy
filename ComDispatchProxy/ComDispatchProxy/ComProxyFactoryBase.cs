using System.Collections.Concurrent;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Xml.Linq;

namespace ComDispatchProxy;

public abstract class ComProxyFactoryBase : IComProxyFactory
{
    private readonly ConcurrentDictionary<Type, Func<object, object, object>> factoryFunctionDictionary = new();

    protected abstract string TargetAssemblyName { get; }

    protected abstract IEnumerable<Type> GetUsedTypes();

    /// <summary>
    /// XML ファイルから対象の COM インターフェイスを読み込み、ProxyCreators を初期化
    /// </summary>
    public ComProxyFactoryBase(string xmlPath)
    {
        factoryFunctionDictionary.Clear();

        var xml = XDocument.Load(xmlPath);

        IEnumerable<string> proxyTargetInterfaces = GetTargetInterfaces(xml);

        IEnumerable<Type> usedTypes = GetUsedTypes();

        CreateFactoryFunctions(proxyTargetInterfaces, usedTypes);

        // -----------------------------------------------------------
        // ローカル関数：ファクトリ関数を生成して、Dictionaryに登録する
        void CreateFactoryFunctions(IEnumerable<string> interfaces, IEnumerable< Type> usedTypes)
        {
            foreach (var type in usedTypes)
            {
                if (type.FullName == null || interfaces.Contains(type.FullName) == false)
                {
                    continue;
                }

                factoryFunctionDictionary[type] = (obj, parentObj) => CreateProxyForType(type, obj, parentObj);
                Debug.WriteLine($"[LOG] Registered proxy for {type.FullName}");
            }
        }
    }

    protected virtual IEnumerable<string> GetTargetInterfaces(XDocument xml)
    {
        return xml
            .Descendants("Generate")
            .Descendants("Assembly")
            .Where(asm => asm.Attribute("name")?.Value == TargetAssemblyName)
            .Descendants("Interface")
            .Select(x => $"{TargetAssemblyName}.{x.Value}");
    }

    /// <summary>
    /// COM オブジェクトの型を判定し、適切なプロキシを生成する。
    ///
    /// 1. `factoryFunctionDictionary` に登録されている型リストを順にチェック。
    /// 2. `IsComObjectOfType()` を使用して `comObject` がその型 (`type`) を実装しているか判定。
    /// 3. 一致する型が見つかれば、その型に対応する `factoryFunction` メソッドを呼び出してプロキシを生成し返す。
    /// 4. 一致する型が見つからなければ、`comObject` をそのまま返す。
    /// </summary>
    /// <param name="comObject">プロキシを作成する COM オブジェクト</param>
    /// <param name="parentObject">親の COM オブジェクト（親子関係を管理するため）</param>
    /// <returns>対応する型の `ComDispatchProxy<T>` インスタンス、または `comObject` そのまま</returns>
    /// <exception cref="ArgumentNullException">`comObject` が `null` の場合</exception>
    public object? CreateProxyByFactoryFunction(object comObject, object parentObject)
    {
        // 1.  `factoryFunctionDictionary` に登録されている型リストを順にチェック。
        foreach (var (type, factoryFunction) in factoryFunctionDictionary)
        {
            // 2. `comObject` が `type` のインターフェースを実装しているか確認
            if (IsComObjectOfType(comObject, type))
            {
                Debug.WriteLine($"[LOG] Creating proxy for type: {type.FullName}");

                // 3. 一致する型が見つかれば、その型に対応する `factoryFunction` メソッドを呼び出してプロキシを生成
                return factoryFunction(comObject, parentObject);
            }
        }

        // 4. 一致する型が見つからなければ、`comObject` をそのまま返す。
        Debug.WriteLine($"[LOG] No matching type found for COM object. Returning original object.");
        return comObject;
    }

    /// <summary>
    /// 指定された COM オブジェクトが、特定のインターフェース (`interfaceType`) を実装しているか確認する。
    /// （COM オブジェクトのGetType()が System.__ComObject しか取得できない問題の解決策）
    ///
    /// 1. `Marshal.GetIUnknownForObject()` を使用して、COM オブジェクトの `IUnknown` ポインタを取得。
    /// 2. `Marshal.QueryInterface()` を使用して、指定された `interfaceType` の GUID に対応するインターフェースを持っているか確認。
    /// 3. `QueryInterface()` が成功した場合 (`S_OK`)、該当インターフェースを実装していると判断し `true` を返す。
    /// 4. 失敗した場合 (`E_NOINTERFACE` など)、その型を実装していないと判断し `false` を返す。
    /// 5. `IUnknown` ポインタと `QueryInterface` で取得したポインタは適切に `Release()` してメモリリークを防ぐ。
    /// </summary>
    /// <param name="comObject">判定対象の COM オブジェクト</param>
    /// <param name="interfaceType">確認したい COM インターフェースの型</param>
    /// <returns>`comObject` が `interfaceType` を実装している場合は `true`、それ以外は `false`</returns>
    private static bool IsComObjectOfType(object comObject, Type interfaceType)
    {
        IntPtr comPointer = IntPtr.Zero; // `IUnknown` のポインタ
        IntPtr queriedPointer = IntPtr.Zero; // `QueryInterface` で取得するポインタ

        try
        {
            // 1. COM オブジェクトの `IUnknown` ポインタを取得
#pragma warning disable CA1416 // OS互換性警告を無視
            comPointer = Marshal.GetIUnknownForObject(comObject);
#pragma warning restore CA1416

            // 2. インターフェースの GUID を取得
            Guid interfaceGuid = interfaceType.GUID;

            // 3. `QueryInterface` を使用して `interfaceType` のインターフェースを持っているか確認
            int hr = Marshal.QueryInterface(comPointer, ref interfaceGuid, out queriedPointer);

            if (hr == 0) // `S_OK` → インターフェースを実装している
            {
                return true;
            }
        }
        finally
        {
            // 4. `QueryInterface` で取得したポインタを解放（成功した場合のみ）
            if (queriedPointer != IntPtr.Zero)
            {
                Marshal.Release(queriedPointer);
            }

            // 5. `IUnknown` のポインタも解放
            if (comPointer != IntPtr.Zero)
            {
                Marshal.Release(comPointer);
            }
        }

        // `QueryInterface` が失敗した場合 (`E_NOINTERFACE` など)、インターフェースを実装していない
        return false;
    }

    /// <summary>
    /// 指定された COM インターフェイス型のプロキシを動的に作成する。
    /// 
    /// 1. `ComDispatchProxy<T>` の `T` を `interfaceType` に設定し、ジェネリック型を動的に生成する。
    /// 2. `CreateProxy` メソッドを取得する。（`CreateProxy(T comObject, object parentObject)`）
    /// 3. `CreateProxy` メソッドを実行し、新しいプロキシインスタンスを作成する。
    /// 4. 作成されたプロキシが `null` の場合は例外をスローする。
    /// </summary>
    /// <param name="interfaceType">プロキシ化する COM インターフェースの型</param>
    /// <param name="comObject">プロキシ化する対象の COM オブジェクト</param>
    /// <param name="parentObject">親となる COM オブジェクト(親オブジェクトのDisposeで子オブジェクトもDisposeするため）</param>
    /// <returns>生成された `ComDispatchProxy<T>` のインスタンス</returns>
    /// <exception cref="InvalidOperationException">メソッドの取得やプロキシ生成に失敗した場合</exception>
    private object CreateProxyForType(Type interfaceType, object comObject, object parentObject)
    {
        // 1. `ComDispatchProxy<T>` のジェネリック型を動的に生成
        var proxyType = typeof(ComDispatchProxy<>).MakeGenericType(interfaceType);

        // 2. `CreateProxy(T comObject, object parentObject)` メソッドを取得
        var createMethod = proxyType.GetMethod("CreateProxy",
            BindingFlags.Static | BindingFlags.Public, // 静的 & 公開メソッドを検索
            null,
            new Type[] { typeof(IComProxyFactory), interfaceType, typeof(object) }, // メソッドの引数型を指定（T, object）
            null);

        // メソッドが見つからない場合は例外をスロー
        if (createMethod == null)
        {
            throw new InvalidOperationException($"Failed to locate 'CreateProxy' method on {proxyType.FullName}");
        }

        // 3. `CreateProxy(T comObject, object parentObject)` を実行してプロキシを作成
        var proxyInstance = createMethod.Invoke(null, new object[] { this, comObject, parentObject });

        // 生成されたプロキシが `null` の場合は例外をスロー
        if (proxyInstance == null)
        {
            throw new InvalidOperationException($"Unhandled COM object type: {proxyType.FullName}");
        }

        // 4. 生成されたプロキシを返す
        return proxyInstance;
    }
}
