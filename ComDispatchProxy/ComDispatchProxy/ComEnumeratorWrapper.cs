using ComDispatchProxy;
using System;
using System.Collections;
using System.Runtime.InteropServices;

internal sealed class ComEnumeratorWrapper : IEnumerator, IDisposable
{
    private readonly IEnumerator _inner;
    private readonly IComDispatchProxy _parent;
    private readonly string _ctxBase;

    public ComEnumeratorWrapper(IEnumerator inner, IComDispatchProxy parent, string ctxBase)
    {
        _inner = inner ?? throw new ArgumentNullException(nameof(inner));
        _parent = parent ?? throw new ArgumentNullException(nameof(parent));
        _ctxBase = ctxBase ?? "";
    }

    public bool MoveNext() => _inner.MoveNext();
    public void Reset() => _inner.Reset();

    public object Current
    {
        get
        {
            var raw = _inner.Current;

            // 既に proxy なら素通し
            if (raw is IComDispatchProxy) return raw;

            if (raw != null && Marshal.IsComObject(raw))
            {
                // COMが返った瞬間＝ここで Track
                _parent.Session.TrackOrReleaseDuplicate(raw, $"{_ctxBase}.Current");

                // 既存のfactory/キャッシュに任せてproxy化
                return _parent.ProxyFactory.CreateProxyByFactoryFunction(raw, _parent);
            }

            return raw!;
        }
    }

    public void Dispose()
    {
        if (_inner is IDisposable d) d.Dispose();
    }
}