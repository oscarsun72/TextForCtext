
using System.Windows.Forms;
using System.Windows.Input;

namespace WindowsFormsApp1
{
    internal static class KeyboardInfo
    {
        /// <summary>
        /// 判斷某鍵是否被按下彈起 
        /// </summary>
        /// <param name="key">要判斷的按鍵</param>
        /// <returns></returns>
        internal static bool getKeyStateToggled(Key key)
        {
            return (Keyboard.GetKeyStates(key) & KeyStates.Toggled) > 0;
        }

        /// <summary>
        /// 判斷某鍵是否沒被按下
        /// </summary>
        /// <param name="key">要判斷的按鍵</param>
        /// <returns></returns>
        internal static bool getKeyStateNone(Key key)
        {
            return (Keyboard.GetKeyStates(key) & KeyStates.None) > 0;
        }

        /// <summary>
        /// 判斷某鍵是否已被按下
        /// </summary>
        /// <param name="key">要判斷的按鍵</param>
        /// <returns></returns>
        internal static bool getKeyStateDown(Key key)
        {
            return (Keyboard.GetKeyStates(key) & KeyStates.Down) > 0;
        }
    }

    // 鍵盤掛鉤管理類
    public static class HookManager
    {
        public static event System.Windows.Forms.KeyEventHandler KeyDown;

        private static void OnKeyDown(System.Windows.Forms.KeyEventArgs e)
        {
            KeyDown?.Invoke(null, e);
        }

        static HookManager()
        {
            // 註冊全域鍵盤掛鉤
            Application.AddMessageFilter(new KeyMessageFilter());
        }

        private class KeyMessageFilter : IMessageFilter
        {
            private const int WM_KEYDOWN = 0x0100;

            public bool PreFilterMessage(ref Message m)
            {
                if (m.Msg == WM_KEYDOWN)
                {
                    Keys keyData = (Keys)m.WParam;
                    System.Windows.Forms.KeyEventArgs e = new System.Windows.Forms.KeyEventArgs(keyData);
                    OnKeyDown(e);
                }
                return false;
            }
        }
    }
}