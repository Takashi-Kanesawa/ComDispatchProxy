using System;
using System.Collections.Generic;
using System.Threading;

namespace ComDispatchProxy
{
    internal static class ComProxyAppDomainHook
    {
        private static int _registered;
        private static readonly object _gate = new();
        private static readonly List<WeakReference<IComDispatchProxy>> _roots = new();

        public static void RegisterRoot(IComDispatchProxy root)
        {
            if (!ComProxyConfig.EnableAppDomainCleanupHandlers) return;

            lock (_gate)
            {
                _roots.Add(new WeakReference<IComDispatchProxy>(root));
                EnsureRegistered();
            }
        }

        private static void EnsureRegistered()
        {
            if (Interlocked.Exchange(ref _registered, 1) == 1) return;

            AppDomain.CurrentDomain.ProcessExit += OnProcessExit;
            AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;
        }

        private static void OnProcessExit(object? sender, EventArgs e) => Cleanup("ProcessExit");
        private static void OnUnhandledException(object? sender, UnhandledExceptionEventArgs e) => Cleanup("UnhandledException");

        private static void Cleanup(string reason)
        {
            lock (_gate)
            {
                for (int i = _roots.Count - 1; i >= 0; i--)
                {
                    if (_roots[i].TryGetTarget(out var root))
                    {
                        try { root.Dispose(); }
                        catch { /* 退出時なので握りつぶし */ }
                    }
                    else
                    {
                        _roots.RemoveAt(i);
                    }
                }
            }
        }
    }
}
