/*----=----=----=----=----=----=----=----=----=----=----=----=----=----=----=----=----
    ████████████   ██          ███████
          ██       ██         ██
          ██       ██          ███████
     ██   ██       ██                ██
      █████        ████████    ███████


    Just License Software (JLS) License
    Copyright (c) 2023 Joshua L Shuller

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE, AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES, OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT, OR OTHERWISE, ARISING FROM,
    OUT OF, OR IN CONNECTION WITH THE SOFTWARE OR THE USE, OR OTHER DEALINGS IN
    THE SOFTWARE.
----=----=----=----=----=----=----=----=----=----=----=----=----=----=----=----=----*/


using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Management;
using System.Reflection;
using System.Resources;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;


namespace ToggleLightDarkMode {


    /// <summary>
    /// A simple program which creates a tray icon to toggle the system and app 
    /// themes to either light or dark mode.
    /// </summary>
    /// <remarks>
    /// This is basically the simplest program I have ever written.
    /// </remarks>
    internal static class Program {


        #region Variables

        /// <summary>
        /// Reference to the tray icon for updating on toggle.
        /// Just easier than casting the "sender" callback parameter to a NotifyIcon.
        /// </summary>
        private static NotifyIcon trayIcon = null;


        /// <summary>
        /// To watch the registry for external changes to theme
        /// </summary>
        private static ManagementEventWatcher registryWatcher = null;


        /// <summary>
        /// A resource manager for getting labels in various languages
        /// </summary>
        private static ResourceManager languages = null;

        #endregion



        #region Main

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        public static void Main(String[] args) {


            // Check if it's being launched from installer
            {
                // This is a dirty workaround for 
                if (args.Length == 1 && string.Equals(args[0], "INSTALLER")) {
                    Process.Start(Application.ExecutablePath); 
                    return; 
                }
            }


            // Get the resource manager for translations first.
            {
                string langCode = System.Globalization.CultureInfo.CurrentUICulture.TwoLetterISOLanguageName;
                if (string.Equals(langCode, "zh")) {
                    langCode = System.Globalization.CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName.ToLower();
                }
                try {
                    languages = new ResourceManager("ToggleLightDarkMode.languages.lang_" + langCode, Assembly.GetExecutingAssembly());
                } catch (Exception x) {
                    languages = null;
                }
                if (languages == null) {
                    try {
                        languages = new ResourceManager("ToggleLightDarkMode.languages.lang_en", Assembly.GetExecutingAssembly());
                    } catch (Exception x) {
                        languages = null;
                        MessageBox.Show("Cannot load language: " + langCode);
                    }
                }
            }


            // Exit if program is already running.
            if (System.Diagnostics.Process.GetProcessesByName(System.IO.Path.GetFileNameWithoutExtension(System.Reflection.Assembly.GetEntryAssembly().Location)).Count() > 1) {
                MessageBox.Show(languages.GetString("program_already_running"));
                return;
            }


            // Attempt to make sure the registryWatcher is properly disposed.
            // Also closes the tray icon.
            // This does not appear to be working when terminated from Task Manager.
            AppDomain.CurrentDomain.ProcessExit += ShutdownCleanupCallback;
            AppDomain.CurrentDomain.UnhandledException += ShutdownCleanupCallback;


            // Start the registry watcher, in case the light/dark mode is changed externally
            StartRegistryWatcher();


            // Check if the current theme is light
            // So that we can display the correct icon in the tray
            int lightMode = GetRegistryDwordValue(
                "HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize",
                "SystemUsesLightTheme",
                -1
            );


            // Show tray icon button according to current theme
            // If the current theme is light, show dark button, otherwise, show light button
            if (lightMode > 0) {
                trayIcon = new NotifyIcon {
                    Icon = Properties.Resources.SetDarkIcon,
                    Text = languages.GetString("switch_to_dark_mode")
                };
            } else {
                trayIcon = new NotifyIcon {
                    Icon = Properties.Resources.SetLightIcon,
                    Text = languages.GetString("switch_to_light_mode")
                };
            }


            // Create a context menu for the system tray icon.
            ContextMenu contextMenu = new ContextMenu();
            contextMenu.MenuItems.Add(languages.GetString("exit"), ExitCallback);


            // Attach the context menu to the NotifyIcon.
            trayIcon.ContextMenu = contextMenu;


            // Add click handler to toggle
            trayIcon.MouseClick += ToggleLightDarkClicked;


            // Show the system tray icon.
            trayIcon.Visible = true;


            // Start a message loop to keep the application running.
            Application.Run();
        }



        /// <summary>
        /// Callback for when the tray icon is clicked
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private static void ToggleLightDarkClicked(object sender, EventArgs e) {

            // Check if the current theme is light
            int lightMode = GetRegistryDwordValue(
                "HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize",
                "SystemUsesLightTheme",
                -1
            );

            // Update current them and tray icon reflecting toggle action
            if (lightMode > 0) {
                trayIcon.Icon = Properties.Resources.SetLightIcon;
                trayIcon.Text = languages.GetString("switch_to_light_mode");
                SetDarkMode();
            } else {
                trayIcon.Icon = Properties.Resources.SetDarkIcon;
                trayIcon.Text = languages.GetString("switch_to_dark_mode");
                SetLightMode();
            }
        }

        #endregion



        #region Helper and Wrapper Methods

        #region App Exit

        /// <summary>
        /// Callback method for when the program shuts down.
        /// This is a counter-measure used for when the program is killed from external sources.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private static void ShutdownCleanupCallback(object sender, EventArgs e) {

            // Close tray icon first, for quicker response.
            // Not so important, but just to be consistent
            trayIcon?.Dispose();
            trayIcon = null;

            // Stop and dispose of the registry watcher
            // If the ManagementEventWatcher is running and the program exits without
            // properly stopping or disposing of the watcher, resources may be wasted,
            // and the watcher might not be cleaned up correctly. 
            try {
                registryWatcher?.Stop();
            } catch (Exception x2) {
            }
            registryWatcher?.Dispose();
            registryWatcher = null;

        }


        /// <summary>
        /// Callback for when "Exit" is clicked (controlled exit)
        /// Exits the application, so it doesn't really matter
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private static void ExitCallback(object sender, EventArgs e) {

            // Hide the tray icon first so that user can see response faster.
            trayIcon.Visible = false;

            // Stop and dispose of the registry watcher
            // If the ManagementEventWatcher is running and the program exits without
            // properly stopping or disposing of the watcher, resources may be wasted,
            // and the watcher might not be cleaned up correctly. 
            try {
                registryWatcher?.Stop();
            } catch (Exception x2) {
            }
            registryWatcher?.Dispose();
            registryWatcher = null;

            // Fix for exit:
            // The click callback is running before the exit callback when exit is clicked, so we have to switch it twice :/
            {
                // Check if the current theme is light
                int lightMode = GetRegistryDwordValue(
                    "HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize",
                    "SystemUsesLightTheme",
                    -1
                );

                // Set the registry to toggled value from which it will revert on exit
                if (lightMode > 0) {
                    SetDarkMode();
                } else {
                    SetLightMode();
                }

            }

            // Dispose of tray icon, because... you know...
            trayIcon?.Dispose();
            trayIcon = null;
            
            // Exit the application.
            Application.Exit();
        }

        #endregion



        #region Registry Watcher

        /// <summary>
        /// Callback to change the tray icon if the registry theme is changed externally.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private static void RegistryThemeSettingsChangedCallback(object sender, EventArgs e) {

            // Tray icon may have already been disposed, in which case - nothing to do
            if (trayIcon == null) {
                return;
            }

            // Check if the current theme is light
            int lightMode = GetRegistryDwordValue(
                "HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize",
                "SystemUsesLightTheme",
                -1
            );

            // Update current them and tray icon reflecting toggle action
            if (lightMode <= 0) {
                trayIcon.Icon = Properties.Resources.SetLightIcon;
                trayIcon.Text = languages.GetString("switch_to_light_mode");
            } else {
                trayIcon.Icon = Properties.Resources.SetDarkIcon;
                trayIcon.Text = languages.GetString("switch_to_dark_mode");
            }
        }


        /// <summary>
        /// Starts the registry watcher and saves if it was successfully started or not
        /// This will allow the tray icon to be updated when the light/dark mode is changed
        /// externally.
        /// </summary>
        private static void StartRegistryWatcher() {
            try {
                //
                // Create WQL query to listen to registry events.
                // This cannot listen on HKEY_CURRENT_USER, though not sure why.
                // Instead, use HKEY_USERS and get the value of the current user value.
                // WQL also needs to escape backslashes, so need to add double: \\\\.
                //
                var currentUser = WindowsIdentity.GetCurrent();
                String registryCallbackQuery = String.Format(
                    "SELECT * FROM RegistryTreeChangeEvent "
                    + " WHERE Hive='HKEY_USERS' "
                    + " AND RootPath='{0}\\\\{1}'",
                    currentUser.User.Value,
                    "SOFTWARE\\\\Microsoft\\\\Windows\\\\CurrentVersion\\\\Themes\\\\Personalize"
                );

                // Start the registry watcher
                registryWatcher = new ManagementEventWatcher(registryCallbackQuery);
                registryWatcher.EventArrived += new EventArrivedEventHandler(RegistryThemeSettingsChangedCallback);
                registryWatcher.Start();

            } catch (Exception x) {
                try {
                    registryWatcher?.Stop();
                } catch (Exception x2) {
                }
                registryWatcher?.Dispose();
                registryWatcher = null;
            }
        }

        #endregion



        #region Toggle Light Dark

        /// <summary>
        /// Sets the registry keys for apps and OS to dark.
        /// </summary>
        private static void SetDarkMode() {
            SetRegistryDwordValue(
                "HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize",
                "AppsUseLightTheme", 
                "0"
            );
            SetRegistryDwordValue(
                "HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize",
                "SystemUsesLightTheme", 
                "0"
            );
        }

        /// <summary>
        /// Sets the registry keys for apps and OS to light.
        /// </summary>
        private static void SetLightMode() {
            SetRegistryDwordValue(
                "HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize", 
                "AppsUseLightTheme", 
                "1"
            );
            SetRegistryDwordValue(
                "HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize", 
                "SystemUsesLightTheme", 
                "1"
            );
            RefreshWindowsExplorer();
            //ToggleDesktopIcons();
            //ToggleDesktopIcons();
        }

        #region Refresh Windows

        [System.Runtime.InteropServices.DllImport("Shell32.dll")]
        private static extern int SHChangeNotify(int eventId, int flags, IntPtr item1, IntPtr item2);


        /// <summary>
        /// Some windows installations do not refresh on update to registry.
        /// </summary>
        private static void RefreshWindowsExplorer() {
            // Refresh the desktop
            SHChangeNotify(0x8000000, 0x1000, IntPtr.Zero, IntPtr.Zero);

            // Refresh any open explorer windows
            // based on http://stackoverflow.com/questions/2488727/refresh-windows-explorer-in-win7
            Guid CLSID_ShellApplication = new Guid("13709620-C279-11CE-A49E-444553540000");
            Type shellApplicationType = Type.GetTypeFromCLSID(CLSID_ShellApplication, true);

            object shellApplication = Activator.CreateInstance(shellApplicationType);
            object windows = shellApplicationType.InvokeMember("Windows", System.Reflection.BindingFlags.InvokeMethod, null, shellApplication, new object[] { });

            Type windowsType = windows.GetType();
            object count = windowsType.InvokeMember("Count", System.Reflection.BindingFlags.GetProperty, null, windows, null);
            for (int i = 0; i < (int)count; i++) {
                object item = windowsType.InvokeMember("Item", System.Reflection.BindingFlags.InvokeMethod, null, windows, new object[] { i });
                Type itemType = item.GetType();

                // Only refresh Windows Explorer, without checking for the name this could refresh open IE windows
                string itemName = (string)itemType.InvokeMember("Name", System.Reflection.BindingFlags.GetProperty, null, item, null);
                if (string.Equals(itemName, "Windows Explorer", StringComparison.OrdinalIgnoreCase)) {
                    itemType.InvokeMember("Refresh", System.Reflection.BindingFlags.InvokeMethod, null, item, null);
                }
            }
        }


    [DllImport("user32.dll", SetLastError = true)]
    static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

    [DllImport("user32.dll", SetLastError = true)]
    static extern IntPtr GetWindow(IntPtr hWnd, GetWindow_Cmd uCmd);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, IntPtr wParam, IntPtr lParam);

    enum GetWindow_Cmd : uint
    {
        GW_HWNDFIRST = 0,
        GW_HWNDLAST = 1,
        GW_HWNDNEXT = 2,
        GW_HWNDPREV = 3,
        GW_OWNER = 4,
        GW_CHILD = 5,
        GW_ENABLEDPOPUP = 6
    }

    private const int WM_COMMAND = 0x111;

    [DllImport("user32.dll", SetLastError = true)]
    static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

    private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder strText, int maxCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowTextLength(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);


    public static string GetWindowText(IntPtr hWnd)
    {
        int size = GetWindowTextLength(hWnd);
        if (size++ > 0)
        {
            var builder = new StringBuilder(size);
            GetWindowText(hWnd, builder, builder.Capacity);
            return builder.ToString();
        }

        return String.Empty;
    }

    public static IEnumerable<IntPtr> FindWindowsWithClass(string className)
    {
        IntPtr found = IntPtr.Zero;
        List<IntPtr> windows = new List<IntPtr>();

        EnumWindows(delegate(IntPtr wnd, IntPtr param)
        {
            StringBuilder cl = new StringBuilder(256);
            GetClassName(wnd, cl, cl.Capacity);
            if (cl.ToString() == className && (GetWindowText(wnd) == "" || GetWindowText(wnd) == null))
            {
                windows.Add(wnd);
            }
            return true;
        },
                    IntPtr.Zero);

        return windows;
    }

    static void ToggleDesktopIcons()
    {
        var toggleDesktopCommand = new IntPtr(0x7402);
        IntPtr hWnd = IntPtr.Zero;
        if (Environment.OSVersion.Version.Major < 6 || Environment.OSVersion.Version.Minor < 2) //7 and -
            hWnd = GetWindow(FindWindow("Progman", "Program Manager"), GetWindow_Cmd.GW_CHILD);
        else
        {
            var ptrs = FindWindowsWithClass("WorkerW");
            int i = 0;
            while (hWnd == IntPtr.Zero && i < ptrs.Count())
            {
                hWnd = FindWindowEx(ptrs.ElementAt(i), IntPtr.Zero, "SHELLDLL_DefView", null);
                i++;
            }
        }
        SendMessage(hWnd, WM_COMMAND, toggleDesktopCommand, IntPtr.Zero);
    }

        #endregion


        /// <summary>
        /// Wrapper method to set registry DWORD value.
        /// </summary>
        /// <param name="keyPath">The registry path</param>
        /// <param name="valueName">The key value name (eg AppsUseLightTheme)</param>
        /// <param name="value">The DWORD (Int32) value to set in the registry</param>
        private static void SetRegistryDwordValue(string keyPath, string valueName, string value) {
            Registry.SetValue(keyPath, valueName, value, RegistryValueKind.DWord);
        }


        /// <summary>
        /// Wrapper method to get registry DWORD value.
        /// </summary>
        /// <param name="keyPath">The registry path</param>
        /// <param name="valueName">The key value name (eg AppsUseLightTheme)</param>
        /// <param name="defaultValue">This value is returned if the key doesn't exist or isn't a DWORD</param>
        /// <returns></returns>
        private static int GetRegistryDwordValue(string keyPath, string valueName, int defaultValue = 0) {
            try {
                object value = Registry.GetValue(keyPath, valueName, defaultValue);

                if (value != null && value is int intValue) {
                    // The value exists and is of type int (DWORD).
                    return intValue;
                } else {
                    // The value does not exist or is not of type int (DWORD).
                    return defaultValue;
                }
            } catch (Exception ex) {
                // Nothing we can do...
                return defaultValue;
            }
        }

        #endregion

        #endregion


    }
}
