using System.Collections;
using System.Runtime.InteropServices;

namespace ComDispatchProxy;

internal sealed class ComEnumeratorWrapper : IEnumerator, IDisposable
{
    private readonly IEnumerator _inner;
    private readonly IComDispatchProxy _parent;
    private readonly string _ctxBase;
    private readonly IDisposable? _innerDisposable;
    private bool _isDisposed;

    public ComEnumeratorWrapper(IEnumerator inner, IComDispatchProxy parent, string ctxBase)
    {
        _inner = inner ?? throw new ArgumentNullException(nameof(inner));
        _parent = parent ?? throw new ArgumentNullException(nameof(parent));
        _ctxBase = ctxBase ?? throw new ArgumentNullException(nameof(ctxBase));
        _innerDisposable = inner as IDisposable;
    }

    public object Current
    {
        get
        {
            ThrowIfDisposed();

            var current = _inner.Current;
            if (current is null)
            {
                return null!;
            }

            if (Marshal.IsComObject(current))
            {
                _parent.Session.TrackOrReleaseDuplicate(current, $"{_ctxBase}.Current");
                return _parent.ProxyFactory.CreateProxyByFactoryFunction(current, _parent);
            }

            return current;
        }
    }

    public bool MoveNext()
    {
        ThrowIfDisposed();
        return _inner.MoveNext();
    }

    public void Reset()
    {
        ThrowIfDisposed();
        _inner.Reset();
    }

    public void Dispose()
    {
        if (_isDisposed)
        {
            return;
        }

        _innerDisposable?.Dispose();
        _isDisposed = true;
    }

    private void ThrowIfDisposed()
    {
        if (_isDisposed)
        {
            throw new ObjectDisposedException(nameof(ComEnumeratorWrapper));
        }

        if (_parent.IsDisposed || _parent.Session.IsDisposed)
        {
            throw new ObjectDisposedException(_parent.ComObjectName);
        }
    }
}
