﻿using Selenium.Core;
using System;
using System.Runtime.InteropServices;
using System.Text;

namespace Selenium {

    /// <summary>
    /// Web driver for Internet Explorer
    /// </summary>
    /// 
    /// <example>
    /// 
    /// VBScript:
    /// <code lang="vbs">	
    /// Class Script
    ///     Dim driver
    ///     
    ///     Sub Class_Initialize
    ///         Set driver = CreateObject("Selenium.IEDriver")
    ///         driver.Get "http://www.google.com"
    ///         ...
    ///     End Sub
    /// 
    ///     Sub Class_Terminate
    ///         driver.Quit
    ///     End Sub
    /// End Class
    /// 
    /// Set s = New Script
    /// </code>
    /// 
    /// VBA:
    /// <code lang="vbs">	
    /// Public Sub Script()
    ///   Dim driver As New IEDriver
    ///   driver.Get "http://www.mywebsite.com"
    ///   ...
    ///   driver.Quit
    /// End Sub
    /// </code>
    /// 
    /// </example>
    [ProgId("Selenium.IEDriver")]
    [Guid("0277FC34-FD1B-4616-BB19-EED04A1E4CD1")]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None)]
    public class IEDriver : WebDriver, ComInterfaces._WebDriver {

        const string BROWSER_NAME = "internet explorer";

        public IEDriver()
            : base(BROWSER_NAME) { }

        internal static IDriverService StartService(WebDriver wd) {
            ExtendCapabilities(wd, false);

            var svc = new DriverService();
            svc.AddArgument("/host=" + svc.IPEndPoint.Address.ToString());
            svc.AddArgument("/port=" + svc.IPEndPoint.Port.ToString());
            svc.AddArgument("/log-level=ERROR");
            svc.AddArgument("/silent");

            string serverName;
            if (wd.Capabilities.TryGetValue("ie.serverBinary", out serverName)) {
                wd.Capabilities.Remove("ie.serverBinary");
            } else {
                serverName = "iedriver.exe";
            }
            svc.Start(serverName, true);
            return svc;
        }

        internal static void ExtendCapabilities(WebDriver wd, bool remote) {
            var capa = wd.Capabilities;
            capa["silent"] = true;
            capa["nativeEvents"] = true;
            capa["initialBrowserUrl"] = "about:blank";
            //capa["ignoreProtectedModeSettings"] = true;
            //capa["requireWindowFocus"] = true;

            if (wd.Arguments.Count > 0) {
                StringBuilder cmd_args = new StringBuilder();
                foreach (string arg in wd.Arguments)
                    cmd_args.Append(arg).Append(' ');

                capa["ie.browserCommandLineSwitches"] = cmd_args.ToString();
                capa["ie.forceCreateProcessApi"] = true;
            }
        }

    }

}
