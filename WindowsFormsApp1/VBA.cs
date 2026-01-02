using word=Microsoft.Office.Interop.Word;
//using Microsoft.Office.Interop.PowerPoint;
using excel=Microsoft.Office.Interop.Excel;
using System;
using System.Reflection;
using WindowsFormsApp1;

namespace TextForCtext
{
    /// <summary>
    /// Copilot大菩薩 20260101元旦 
    /// </summary>
    public static class MacroRunner
    {//https://copilot.microsoft.com/shares/dC9VeQV1fgytFSk5odKq7
     //https://copilot.microsoft.com/shares/iyEtcu4WH3qhe688DBg9e

        /// <summary>
        /// 呼叫 Word VBA 巨集
        /// 只要呼叫 MacroRunner.Run(...) 就能統一處理「無引數、有引數、不同引數數量」的情境。
        /// </summary>
        /// <param name="app">Word Application 物件</param>
        /// <param name="macroName">巨集名稱 (可含模組名，如 "Module1.MyMacro")</param>
        /// <param name="args">可變引數，依巨集定義傳入</param>
        public static void Run(word.Application app, string macroName, params object[] args)
        {
            try
            {
                if (app == null)
                    throw new ArgumentNullException(nameof(app), "Word Application 尚未初始化");

                if (string.IsNullOrEmpty(macroName))
                    throw new ArgumentException("巨集名稱不可為空");

                if (args == null)
                    args = new object[0];

                // 自動型別修正：將 int 轉成 byte 或 long
                for (int i = 0; i < args.Length; i++)
                {
                    if (args[i] is int intVal)
                    {
                        if (intVal >= 0 && intVal <= 255)
                            args[i] = (byte)intVal;
                        else
                            args[i] = (long)intVal;
                    }
                }

                dynamic appDyn = app;//https://copilot.microsoft.com/shares/q9VmgTJduHYdwiAgRxhho 2026年元旦 https://copilot.microsoft.com/shares/673wiBgfzujc4j5yr2Lrt
                                     //https://copilot.microsoft.com/shares/biDNcWYyWYhtsF2kzJQw7

                // 展開最多 30 個參數
                switch (args.Length)
                {
                    case 0: appDyn.Run(macroName); break;
                    case 1: appDyn.Run(macroName, args[0]); break;
                    case 2: appDyn.Run(macroName, args[0], args[1]); break;
                    case 3: appDyn.Run(macroName, args[0], args[1], args[2]); break;
                    case 4: appDyn.Run(macroName, args[0], args[1], args[2], args[3]); break;
                    case 5: appDyn.Run(macroName, args[0], args[1], args[2], args[3], args[4]); break;
                    // …依序展開到 30
                    case 30:
                        appDyn.Run(macroName,
                        args[0], args[1], args[2], args[3], args[4],
                        args[5], args[6], args[7], args[8], args[9],
                        args[10], args[11], args[12], args[13], args[14],
                        args[15], args[16], args[17], args[18], args[19],
                        args[20], args[21], args[22], args[23], args[24],
                        args[25], args[26], args[27], args[28], args[29]); break;
                    default:
                        throw new ArgumentException("Word 巨集最多支援 30 個引數");
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine($"呼叫巨集 {macroName} 時發生錯誤：{ex.Message}");
            }
        }
        //public static void Run(Application app, string macroName, params object[] args)
        //{
        //    try
        //    {
        //        if (args == null || args.Length == 0)
        //        {
        //            app.Run(macroName);
        //        }
        //        else
        //        {
        //            app.Run(macroName, args);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(
        //            $"呼叫巨集 {macroName} 時發生錯誤：{ex.Message}",
        //            "巨集呼叫錯誤");
        //        Console.WriteLine($"呼叫巨集 {macroName} 時發生錯誤：{ex.Message}");
        //    }
        //}
    }



    /** Project Structure: https://copilot.microsoft.com/shares/DQJFj8C78BFsLfATq4VhT

    OfficeMacroInvoker/
 ├── IOfficeMacroInvoker.cs   # 統一介面
 ├── WordMacroInvoker.cs      # Word 實作
 ├── ExcelMacroInvoker.cs     # Excel 實作
 ├── PowerPointMacroInvoker.cs# PowerPoint 實作
 ├── MacroUtils.cs            # 共用工具類別
 └── Tests/                   # 單元測試範例
    **/
    public interface IOfficeMacroInvoker
    {
        void RunMacro(string macroName, params object[] args);
    }


    public class WordMacroInvoker : IOfficeMacroInvoker
    {
        private readonly word.Application _app;

        public WordMacroInvoker(word.Application app)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
        }

        public void RunMacro(string macroName, params object[] args)
        {
            MacroUtils.NormalizeArgs(ref args);
            dynamic appDyn = _app;
            MacroUtils.DispatchRun(appDyn, macroName, args);
        }
    }



    public class ExcelMacroInvoker : IOfficeMacroInvoker
    {
        private readonly excel.Application _app;

        public ExcelMacroInvoker(excel.Application app)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
        }

        public void RunMacro(string macroName, params object[] args)
        {
            MacroUtils.NormalizeArgs(ref args);
            dynamic appDyn = _app;
            MacroUtils.DispatchRun(appDyn, macroName, args);
        }
    }


    //public class PowerPointMacroInvoker : IOfficeMacroInvoker
    //    {
    //        private readonly Application _app;

    //        public PowerPointMacroInvoker(Application app)
    //        {
    //            _app = app ?? throw new ArgumentNullException(nameof(app));
    //        }

    //        public void RunMacro(string macroName, params object[] args)
    //        {
    //            MacroUtils.NormalizeArgs(ref args);
    //            dynamic appDyn = _app;
    //            MacroUtils.DispatchRun(appDyn, macroName, args);
    //        }
    //    }

    public static class MacroUtils
    {
        /// <summary>
        /// 取得空的 object[]，兼容新舊 .NET
        /// </summary>
        public static object[] EmptyArgs()
        {
#if NET46_OR_GREATER || NETCOREAPP
                                return Array.Empty<object>();
#else
            return new object[0];
#endif
        }

        public static void NormalizeArgs(ref object[] args)
        {
            if (args == null) { args = new object[0]; return; }

            for (int i = 0; i < args.Length; i++)
            {
                if (args[i] is int intVal)
                {
                    if (intVal >= 0 && intVal <= 255)
                        args[i] = (byte)intVal;
                    else
                        args[i] = (long)intVal;
                }
            }
        }

        public static void DispatchRun(dynamic appDyn, string macroName, object[] args)
        {
            switch (args.Length)
            {
                case 0: appDyn.Run(macroName); break;
                case 1: appDyn.Run(macroName, args[0]); break;
                case 2: appDyn.Run(macroName, args[0], args[1]); break;
                case 3: appDyn.Run(macroName, args[0], args[1], args[2]); break;
                // …依序展開到 30
                case 30:
                    appDyn.Run(macroName,
                    args[0], args[1], args[2], args[3], args[4],
                    args[5], args[6], args[7], args[8], args[9],
                    args[10], args[11], args[12], args[13], args[14],
                    args[15], args[16], args[17], args[18], args[19],
                    args[20], args[21], args[22], args[23], args[24],
                    args[25], args[26], args[27], args[28], args[29]); break;
                default:
                    throw new ArgumentException("Office 巨集最多支援 30 個引數");
            }
        }
    }


}