using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;

namespace ComDispatchProxy;

/// <summary>
/// Root の using スコープに紐づく、COM Proxy 管理の「セッション」。
/// - Root/Parent 参照は持たない（相互参照を避ける）
/// - RCW → Proxy キャッシュだけを持つ（Proxy は弱参照で保持）
/// </summary>
internal sealed class ComProxySession : IDisposable
{
    private readonly object _gate = new();
    private bool _disposed;

    // interfaceType ごとに RCW(参照等価) → Proxy(弱参照) をキャッシュ
    private readonly Dictionary<Type, Dictionary<object, WeakReference<object>>> _byInterface = new();

    public bool IsDisposed
    {
        get { lock (_gate) return _disposed; }
    }

    public bool TryGet(Type interfaceType, object rcw, out object proxy)
    {
        lock (_gate)
        {
            if (_disposed)
            {
                proxy = null!;
                return false;
            }

            if (_byInterface.TryGetValue(interfaceType, out var map) &&
                map.TryGetValue(rcw, out var weak) &&
                weak.TryGetTarget(out proxy!))
            {
                // 既に Release 済みならキャッシュから除去して再生成
                if (proxy is IComDispatchProxy dp && dp.WasReleased)
                {
                    map.Remove(rcw);
                    proxy = null!;
                    return false;
                }

                ComProxyLog.Write($"[ComDispatchProxy CACHE HIT] {interfaceType.FullName} RCW={RuntimeHelpers.GetHashCode(rcw)} Proxy={RuntimeHelpers.GetHashCode(proxy)}");
                return true;
            }

            proxy = null!;
            return false;
        }
    }

    public void Put(Type interfaceType, object rcw, object proxy)
    {
        lock (_gate)
        {
            if (_disposed) return;

            if (!_byInterface.TryGetValue(interfaceType, out var map))
            {
                map = new Dictionary<object, WeakReference<object>>(ReferenceEqualityComparer.Instance);
                _byInterface[interfaceType] = map;
            }

            ComProxyLog.Write($"[ComDispatchProxy CACHE PUT] {interfaceType.FullName} RCW={RuntimeHelpers.GetHashCode(rcw)} Proxy={RuntimeHelpers.GetHashCode(proxy)}");
            map[rcw] = new WeakReference<object>(proxy);
        }
    }

    public void Dispose()
    {
        lock (_gate)
        {
            if (_disposed) return;
            _byInterface.Clear();
            _disposed = true;
        }
    }

    private sealed class ReferenceEqualityComparer : IEqualityComparer<object>
    {
        public static readonly ReferenceEqualityComparer Instance = new();
        public new bool Equals(object? x, object? y) => ReferenceEquals(x, y);
        public int GetHashCode(object obj) => RuntimeHelpers.GetHashCode(obj);
    }
}

internal interface IComProxySessionProvider
{
    ComProxySession Session { get; }
}
