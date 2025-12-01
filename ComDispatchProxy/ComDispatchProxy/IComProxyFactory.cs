namespace ComDispatchProxy;

/// <summary>
/// COMオブジェクトのプロキシを作成するファクトリのインターフェース
/// </summary>
public interface IComProxyFactory
{
    /// <summary>
    /// COM オブジェクトの型を判定し、適切なプロキシを生成する。
    /// </summary>
    /// <param name="comObject">プロキシを作成する COM オブジェクト</param>
    /// <param name="parentObject">親の COM オブジェクト</param>
    /// <returns>対応する型のプロキシインスタンス、または `comObject` そのまま</returns>
    object? CreateProxyByFactoryFunction(object comObject, object parentObject);
}
