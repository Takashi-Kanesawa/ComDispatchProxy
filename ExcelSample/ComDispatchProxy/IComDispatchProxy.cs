using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ComDispatchProxy;

#region IComDispatchProxy インターフェイス定義
/// <summary>
/// 生のCOMオブジェクトを取得するためのインターフェイス
/// </summary>
public interface IComDispatchProxy : IDisposable
{
    /// <summary>
    /// Dispose済み
    /// </summary>
    bool WasReleased { get; }

    /// <summary>
    /// 生のCOMオブジェクトを取得します
    /// </summary>
    object? RowObject { get; }

    IComDispatchProxy? ParentProxy { get; }

    void AddChild(IComDispatchProxy childObject);

    void RemoveChild(IComDispatchProxy childObject);
}
#endregion
