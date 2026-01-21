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

        // Enabled=true の場合でも StackTrace の取得を抑止できるようにする
        // false の場合は Environment.StackTrace を一切呼ばない
        public static bool CaptureStackTrace { get; set; } = false;
    }
}
