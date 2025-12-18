namespace ComDispatchProxy;

#region IComDispatchProxy インターフェイス定義
/// <summary>
/// 生のCOMオブジェクトを取得するためのインターフェイス
/// </summary>
public interface IComDispatchProxy : IDisposable
{
    /// <summary>
    /// 
    /// </summary>
    IComProxyFactory ProxyFactory { get; }

    /// <summary>
    /// COMオブジェクトをRelease済み
    /// </summary>
    bool WasReleased { get; }

    /// <summary>
    /// COMのルート（最上位の親オブジェクト）になるオブジェクトかを判定
    /// </summary>
    bool IsRoot { get; }

    /// <summary>
    /// 生のCOMオブジェクトを取得します
    /// </summary>
    internal object? RowObject { get; }

    void AddChild(IComDispatchProxy childObject);

    void RemoveChild(IComDispatchProxy childObject);
}
#endregion
