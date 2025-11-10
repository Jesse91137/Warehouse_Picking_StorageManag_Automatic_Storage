
namespace Automatic_Storage.Utilities
{
    /// <summary>
    /// 提供與原生 Win32 API 互動的輔助方法。
    /// </summary>
    public static class NativeHelpers
    {
        /// <summary>
        /// 釋放目前滑鼠捕獲，允許其他視窗接收滑鼠輸入。
        /// </summary>
        /// <remarks>
        /// 此方法會呼叫 user32.dll 的 ReleaseCapture 函式，常用於自訂視窗拖曳等情境。
        /// </remarks>
        /// <returns>
        /// 如果成功則回傳 <c>true</c>，否則回傳 <c>false</c>。
        /// </returns>
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern bool ReleaseCapture();
    }
}
