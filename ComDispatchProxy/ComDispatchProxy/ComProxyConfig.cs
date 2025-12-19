namespace ComDispatchProxy
{
    public static class ComProxyConfig
    {
        /// <summary>
        /// true の場合、ComDispatchProxy.Dispose() で Marshal.ReleaseComObject を呼ぶ（強制解放）。
        /// false の場合、ReleaseComObject は呼ばない（解放タイミングはRCW/GC側に委ねる）。
        /// ※このフラグ自体は1番ではまだ参照されません。次のステップで Dispose 側に反映します。
        /// </summary>
        public static bool AggressiveReleaseComObjects { get; set; } = true;

        public static bool EnableAppDomainCleanupHandlers { get; set; } = true;

    }
}
