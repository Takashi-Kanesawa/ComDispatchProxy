using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace ComDispatchProxy;

/// <summary>
/// Root の using スコープに紐づく、COM Proxy 管理の「セッション」。
/// - Root/Parent 参照は持たない（相互参照を避ける）
/// - RCW → Proxy キャッシュだけを持つ（Proxy は弱参照で保持）
/// </summary>
internal sealed class ComProxySession : IDisposable
{
    private readonly object _gate = new();
    private bool _isDisposed;

    // interfaceType ごとに RCW(参照等価) → Proxy(弱参照) をキャッシュ
    Dictionary<object, WeakReference<IComDispatchProxy>> _proxyCache = null!;


    // 追加：RCW の Track（ユニーク + 順序 + 取得元）
    private readonly HashSet<object> _trackedRcws;
    private readonly List<object> _trackOrder;
    private readonly Dictionary<object, OriginInfo> _originByRcw;

    private sealed class OriginInfo
    {
        public string Context { get; }
        public string? StackTrace { get; }
        public OriginInfo(string context, string? stackTrace)
        {
            Context = context;
            StackTrace = stackTrace;
        }
    }

    public bool IsDisposed => this._isDisposed;

    public bool AggressiveReleaseComObjects { get; }

    public ComProxySession(bool aggressiveReleaseComObjects)
    {
        this.AggressiveReleaseComObjects = aggressiveReleaseComObjects;
        this._proxyCache = new Dictionary<object, WeakReference<IComDispatchProxy>>(ReferenceEqualityComparer.Instance);
        this._trackedRcws = new HashSet<object>(ReferenceEqualityComparer.Instance);
        this._trackOrder = new List<object>();
        this._originByRcw = new Dictionary<object, OriginInfo>(ReferenceEqualityComparer.Instance);
    }

    /// <summary>
    /// 戻り値COMを Track。すでに同じRCWが戻ってきた場合は、その場でReleaseComObject(1回)して増分を相殺する。
    /// out/ref や COM配列は方針として扱わない。
    /// </summary>
    internal bool TrackOrReleaseDuplicate(object rcw, string context)
    {
        if (this.IsDisposed) return false;
        if (rcw is null) return false;
        if (!Marshal.IsComObject(rcw)) return false;

        if (_trackedRcws.Add(rcw))
        {
            _trackOrder.Add(rcw);
            string? stack = null;
            if (ComProxyLogConfig.Enabled && ComProxyLogConfig.CaptureStackTrace)
            {
                stack = FilterStackTrace(Environment.StackTrace); // ★ここだけが高コスト
            }
            var origin = new OriginInfo(context, stack);
            this._originByRcw[rcw] = origin;

            if (ComProxyLogConfig.Enabled)
            {
                if (origin.StackTrace is { Length: > 0 })
                    ComProxyLog.Write($"[TRACK #{_trackOrder.Count - 1}] ctx={context} rcwHash={RuntimeHelpers.GetHashCode(rcw)}\n{origin.StackTrace}");
                else
                    ComProxyLog.Write($"[TRACK #{_trackOrder.Count - 1}] ctx={context} rcwHash={RuntimeHelpers.GetHashCode(rcw)}");
            }

            return true;
        }

        // duplicate: 参照カウント増分をその場で相殺（Release 1回）
        _originByRcw.TryGetValue(rcw, out var first);
#pragma warning disable CA1416 // OS互換性警告を無視
        int remain = Marshal.ReleaseComObject(rcw);
#pragma warning restore CA1416

        if (ComProxyLogConfig.Enabled)
        {
            if (first?.StackTrace is { Length: > 0 } st)
                ComProxyLog.Write($"[DUP-RELEASE] remain={remain} nowCtx={context} firstCtx={first?.Context ?? "<missing>"} rcwHash={RuntimeHelpers.GetHashCode(rcw)}\n{st}");
            else
                ComProxyLog.Write($"[DUP-RELEASE] remain={remain} nowCtx={context} firstCtx={first?.Context ?? "<missing>"} rcwHash={RuntimeHelpers.GetHashCode(rcw)}");
        }

        // remain==0 は「想定より落ちた」シグナル。ここは fail-fast で安全側に倒す。
        // （必要なら後で設定で抑制できるようにする）
        if (remain == 0)
            throw new InvalidOperationException(
                $"Duplicate RCW release reached 0. nowCtx={context}, firstCtx={first?.Context ?? "<missing>"}");

        return false;
    }


    public void Dispose()
    {
        lock (_gate)
        {
            if (_isDisposed)
            {
                return;
            }

            _isDisposed = true;


            if (AggressiveReleaseComObjects)
            {
                this.ReleaseAllComObjectsInSession();
            }

            _originByRcw.Clear();
            _trackOrder.Clear();
            _trackedRcws.Clear();
            _proxyCache.Clear();
        }
    }

    private void ReleaseAllComObjectsInSession()
    {
        for (int i = _trackOrder.Count - 1; i >= 0; i--)
        {
            var rcw = _trackOrder[i];
            try
            {
                if (!Marshal.IsComObject(rcw)) continue;
                _originByRcw.TryGetValue(rcw, out var origin);
#pragma warning disable CA1416 // OS互換性警告を無視
                int remain = Marshal.ReleaseComObject(rcw); // ユニークRCWに対し1回だけ
#pragma warning restore CA1416

                if (ComProxyLogConfig.Enabled)
                {
                    if (origin?.StackTrace is { Length: > 0 } st)
                        ComProxyLog.Write($"[RELEASE #{i}] remain={remain} firstCtx={origin?.Context ?? "<missing>"} rcwHash={RuntimeHelpers.GetHashCode(rcw)}\n{st}");
                    else
                        ComProxyLog.Write($"[RELEASE #{i}] remain={remain} firstCtx={origin?.Context ?? "<missing>"} rcwHash={RuntimeHelpers.GetHashCode(rcw)}");
                }
            }
            catch (Exception ex)
            {
                ComProxyLog.Write($"[RELEASE FAILED #{i}] {ex.GetType().Name}: {ex.Message}");
            }
        }
    }

    #region ヘルパ

    // スタックトレースのフィルタリングメソッド
    private static string FilterStackTrace(string stackTrace)
    {
        Regex[] includesRegex = { new Regex(@"cs:line \d+$") };

        // スタックトレースの行ごとにフィルタリング
        var filteredStackTrace = string.Join(Environment.NewLine, stackTrace
            .Split(new[] { Environment.NewLine }, StringSplitOptions.None)
            .Where(line =>
                includesRegex.Any(regex => regex.IsMatch(line)) &&
                line.StartsWith("   at ComDispatchProxy`") == false));

        return filteredStackTrace;
    }
    #endregion


    private sealed class ReferenceEqualityComparer : IEqualityComparer<object>
    {
        public static readonly ReferenceEqualityComparer Instance = new();
        public new bool Equals(object? x, object? y) => ReferenceEquals(x, y);
        public int GetHashCode(object obj) => RuntimeHelpers.GetHashCode(obj);
    }

}

public interface IComProxySessionProvider
{
    internal ComProxySession Session { get; }
}
