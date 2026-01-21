using System.Diagnostics;

namespace ComDispatchProxy
{
    internal static class ComProxyLog
    {
        // 全体の ON/OFF スイッチ
        internal static bool Enabled { get; set; } = false;

        // ログ出力先（差し替え可能）
        // null なら ComProxyLog.Write にフォールバック
        internal static Action<string>? Writer { get; set; }

        internal static void Write(string message)
        {
            if (!Enabled)
            {
                return;
            }

            message = $"[ComDispatchProxy LOG] {message}";

            if (Writer != null)
            {
                Writer(message);
            }
            else
            {
#if DEBUG
                Debug.WriteLine(message);
#endif
            }
        }
    }
}
