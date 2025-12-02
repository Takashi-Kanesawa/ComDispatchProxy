namespace ComDispatchProxy
{
    public static class ComProxyLogConfig
    {
        public static bool Enabled
        {
            get => ComProxyLog.Enabled;
            set => ComProxyLog.Enabled = value;
        }

        public static Action<string>? Writer
        {
            get => ComProxyLog.Writer;
            set => ComProxyLog.Writer = value;
        }
    }
}
