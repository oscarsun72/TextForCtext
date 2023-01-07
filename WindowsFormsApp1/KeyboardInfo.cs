
//using System.Windows.Forms;
using System.Windows.Input;

namespace WindowsFormsApp1
{
    internal class KeyboardInfo
    {
        //判斷某鍵是否被按下彈起 
        internal static bool getKeyStateToggled(Key key)
        {
            //return Keyboard.GetKeyStates(key) == KeyStates.Toggled;
            return (Keyboard.GetKeyStates(key) & KeyStates.Toggled) > 0;
            //(Keyboard.GetKeyStates(Key.Return) & KeyStates.Down) > 0//https://learn.microsoft.com/zh-tw/dotnet/api/system.windows.input.keyboard.getkeystates?view=netframework-4.5.2#system-windows-input-keyboard-getkeystates(system-windows-input-key)
        }
        internal static bool getKeyStateNone(Key key)
        {
            return (Keyboard.GetKeyStates(key) & KeyStates.None) > 0;
        }
        internal static bool getKeyStateDown(Key key)
        {
            return (Keyboard.GetKeyStates(key) & KeyStates.Down) > 0;
        }
    }
}