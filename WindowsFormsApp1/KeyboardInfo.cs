
//using System.Windows.Forms;
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
            //return Keyboard.GetKeyStates(key) == KeyStates.Toggled;
            return (Keyboard.GetKeyStates(key) & KeyStates.Toggled) > 0;
            //(Keyboard.GetKeyStates(Key.Return) & KeyStates.Down) > 0//https://learn.microsoft.com/zh-tw/dotnet/api/system.windows.input.keyboard.getkeystates?view=netframework-4.5.2#system-windows-input-keyboard-getkeystates(system-windows-input-key)
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
}