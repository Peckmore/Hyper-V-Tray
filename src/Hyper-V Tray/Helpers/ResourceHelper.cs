using HyperVTray.Resources;
using Microsoft.Win32;
using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Resources;

namespace HyperVTray.Helpers
{
    /// <summary>
    /// Provides quick access to resources that can come from either our local resources, or straight from Hyper-V.
    /// </summary>
    /// <remarks>
    /// <para>In order to ease our localisation burden, we try and use resources directly from the Hyper-V Management Tools, which means
    /// most of our strings are translated for us, and we also have greater consistency with the official tools.</para>
    /// <para>This class acts as a wrapper and will try to get the appropriate resource directly from the tools but, if the resource
    /// can't be loaded, will grab the fallback value from our own embedded resources.</para>
    /// <para>There are also some resources which have no equivalent in the official tools, so for those we have our own resources.</para>
    /// </remarks>
    internal static class ResourceHelper
    {
        #region Fields

        private static ResourceManager? _clientResourceManager;
        private static ResourceManager? _vmBrowserResourceManager;

        #endregion

        #region Events

        public static event EventHandler? ThemeChanged;
        #endregion
        
        #region Construction

        static ResourceHelper()
        {
            LoadIcons();

            // We'll hook into the event for user preference changes, to update the icon if the user changes light/dark mode.
            SystemEvents.UserPreferenceChanged += SystemEvents_UserPreferenceChanged;
        }

        #endregion

        #region Properties

        public static string Button_DontReset => GetStringResource(_clientResourceManager, "ConfirmationResetButton2", StringsFallback.Button_DontReset);
        public static string Button_DontShutDown => GetStringResource(_clientResourceManager, "ConfirmationShutdownButton2", StringsFallback.Button_DontShutDown);
        public static string Button_DontTurnOff => GetStringResource(_clientResourceManager, "ConfirmationTurnoffButton2", StringsFallback.Button_DontTurnOff);
        public static string Button_Reset => GetStringResource(_clientResourceManager, "ConfirmationResetButton1", StringsFallback.Button_Reset);
        public static string Button_ShutDown => GetStringResource(_clientResourceManager, "ConfirmationShutdownButton1", StringsFallback.Button_ShutDown);
        public static string Button_TurnOff => GetStringResource(_clientResourceManager, "ConfirmationTurnoffButton1", StringsFallback.Button_TurnOff);
        public static string Command_Connect => GetStringResource(_vmBrowserResourceManager, "VMOpen_Name", StringsFallback.Command_Connect);
        public static string Command_HyperVManager => GetStringResource(_vmBrowserResourceManager, "SnapInNode_DisplayName", StringsFallback.Menu_HyperVManager);
        public static string Command_Pause => GetStringResource(_vmBrowserResourceManager, "VMPause_Name", StringsFallback.Command_Pause);
        public static string Command_Reset => GetStringResource(_vmBrowserResourceManager, "VMReset_Name", StringsFallback.Command_Reset);
        public static string Command_Resume => GetStringResource(_vmBrowserResourceManager, "VMResume_Name", StringsFallback.Command_Resume);
        public static string Command_Save => GetStringResource(_vmBrowserResourceManager, "VMSaveState_Name", StringsFallback.Command_Save);
        public static string Command_ShutDown => GetStringResource(_vmBrowserResourceManager, "VMShutDown_Name", StringsFallback.Command_ShutDown);
        public static string Command_Start => GetStringResource(_vmBrowserResourceManager, "VMStart_Name", StringsFallback.Command_Start);
        public static string Command_TurnOff => GetStringResource(_vmBrowserResourceManager, "VMTurnOff_Name", StringsFallback.Command_TurnOff);
        public static Icon Icon_HyperV_Critical { get; private set; }
        public static Icon Icon_HyperV_Off { get; private set; }
        public static Icon Icon_HyperV_Running { get; private set; }
        public static string Menu_AllVirtualMachines => Strings.Menu_AllVirtualMachines;
        public static string Message_ConfirmationReset => GetStringResource(_clientResourceManager, "ConfirmationReset", StringsFallback.Message_ConfirmationReset);
        public static string Message_ConfirmationResetMultiple => GetStringResource(_clientResourceManager, "ConfirmationResetMultiple", StringsFallback.Message_ConfirmationResetMultiple);
        public static string Message_ConfirmationShutDown => GetStringResource(_clientResourceManager, "ConfirmationShutdown", StringsFallback.Message_ConfirmationShutDown);
        public static string Message_ConfirmationShutDownMultiple => GetStringResource(_clientResourceManager, "ConfirmationShutdownMultiple", StringsFallback.Message_ConfirmationShutDownMultiple);
        public static string Message_ConfirmationTurnOff => GetStringResource(_clientResourceManager, "ConfirmationTurnoff", StringsFallback.Message_ConfirmationTurnOff);
        public static string Message_ConfirmationTurnOffMultiple => GetStringResource(_clientResourceManager, "ConfirmationTurnoffMultiple", StringsFallback.Message_ConfirmationTurnOffMultiple);
        public static string Message_OpenHyperVManagerFailed => Strings.Message_OpenHyperVManagerFailed;
        public static string Message_OpenVMConnectFailed => GetStringResource(_vmBrowserResourceManager, "Message_OpenVMConnectFailed", StringsFallback.Message_OpenVMConnectFailed);
        public static string Message_PauseVMFailed => GetStringResource(_vmBrowserResourceManager, "Message_PauseVMFailed_Format", StringsFallback.Message_PauseVMFailed);
        public static string Message_PauseVMFailedMultiple => GetStringResource(_vmBrowserResourceManager, "Message_PauseVMFailed", StringsFallback.Message_PauseVMFailedMultiple);
        public static string Message_PowerOffVMFailed => GetStringResource(_vmBrowserResourceManager, "Message_PowerOffFailed_Format", StringsFallback.Message_PowerOffVMFailed);
        public static string Message_PowerOffVMFailedMultiple => GetStringResource(_vmBrowserResourceManager, "Message_PowerOffVMFailed", StringsFallback.Message_PowerOffVMFailed);
        public static string Message_ResetVMFailed => GetStringResource(_vmBrowserResourceManager, "Message_ResetVMFailed_Format", StringsFallback.Message_ResetVMFailed);
        public static string Message_ResetVMFailedMultiple => GetStringResource(_vmBrowserResourceManager, "Message_ResetVMFailed", StringsFallback.Message_ResetVMFailedMultiple);
        public static string Message_ResumeVMFailed => GetStringResource(_vmBrowserResourceManager, "Message_ResumeVMFailed", StringsFallback.Message_ResumeVMFailedMultiple);
        public static string Message_ResumeVMFailedMultiple => GetStringResource(_vmBrowserResourceManager, "Message_ResumeVMFailed", StringsFallback.Message_ResumeVMFailedMultiple);
        public static string Message_SaveStateVMFailed => GetStringResource(_vmBrowserResourceManager, "Message_SaveStateVMFailed_Format", StringsFallback.Message_SaveStateVMFailed);
        public static string Message_SaveStateVMFailedMultiple => GetStringResource(_vmBrowserResourceManager, "Message_SaveStateVMFailed", StringsFallback.Message_SaveStateVMFailedMultiple);
        public static string Message_ShutDownVMFailed => GetStringResource(_vmBrowserResourceManager, "Message_ShutDownVMFailed", StringsFallback.Message_ShutDownVMFailedMultiple);
        public static string Message_ShutDownVMFailedMultiple => GetStringResource(_vmBrowserResourceManager, "Message_ShutDownVMFailed", StringsFallback.Message_ShutDownVMFailedMultiple);
        public static string Message_StartVMFailed => GetStringResource(_vmBrowserResourceManager, "Message_StartVMFailed_Format", StringsFallback.Message_StartVMFailed);
        public static string Message_StartVMFailedMultiple => GetStringResource(_vmBrowserResourceManager, "Message_StartVMFailed", StringsFallback.Message_StartVMFailedMultiple);
        public static string State_Critical => GetStringResource(_vmBrowserResourceManager, "VMOnFailureStateFormat", StringsFallback.State_Critical);
        public static string State_Off => GetStringResource(_vmBrowserResourceManager, "DDDDDD", StringsFallback.State_Off);
        public static string State_Paused => GetStringResource(_vmBrowserResourceManager, "DDDDDD", StringsFallback.State_Paused);
        public static string State_Running => GetStringResource(_vmBrowserResourceManager, "rrrrrrr", StringsFallback.State_Running);
        public static string State_Saved => GetStringResource(_vmBrowserResourceManager, "dddddddd", StringsFallback.State_Saved);
        public static string State_Unknown => GetStringResource(_vmBrowserResourceManager, "dddddddd", StringsFallback.State_Unknown);
        public static string Title_ResetMachine => GetStringResource(_clientResourceManager, "ConfirmationResetTitle", StringsFallback.Title_ResetMachine);
        public static string Title_ShutDownMachine => GetStringResource(_clientResourceManager, "ConfirmationShutdownTitle", StringsFallback.Title_ShutDownMachine);
        public static string Title_TurnOffMachine => GetStringResource(_clientResourceManager, "ConfirmationTurnoffTitle", StringsFallback.Title_TurnOffMachine);
        public static string Toast_CriticalState => Strings.Toast_CriticalState;
        public static string String_UnknownVirtualMachine => Strings.String_UnknownVirtualMachine;

        #endregion

        #region Methods

        #region Event Handlers

        private static void SystemEvents_UserPreferenceChanged(object sender, UserPreferenceChangedEventArgs e)
        {
            if (e.Category == UserPreferenceCategory.General)
            {
                // The user may have changed their theme preferences, so reload our icons (light/dark mode).
                LoadIcons();

                // Notify listeners that the theme (and therefore icons) have changed.
                ThemeChanged?.Invoke(null, EventArgs.Empty);
            }
        }

        #endregion

        #region Private Static

        private static Icon GetIconResource(string filename)
        {
            // Grab an icon resource from our embedded resources, and return it as an `Icon`.

            using (var resStream = typeof(ResourceHelper).Assembly.GetManifestResourceStream($"HyperVTray.Resources.Icons.{filename}.ico"))
            {
                return new Icon(resStream!);
            }
        }
        private static string GetStringResource(ResourceManager? resourceManager, string resourceName, string fallbackValue)
        {
            // Attempt to get a string from the Hyper-V resources, and return the fallback value if the resource could not be found.
            return resourceManager?.GetString(resourceName) ?? fallbackValue;
        }
        private static void LoadIcons()
        {
            bool isDarkMode;
             
            // If the WebView2 profile is set to auto, we'll check the registry to see whether the app should be in light or dark mode.
            using (var themeRegistryKey = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"))
            {
                isDarkMode = (themeRegistryKey?.GetValue("AppsUseLightTheme") as int? ?? 1) == 0;
            }

            // Determine the theme to identify which icons to load.
            var theme = isDarkMode ? "Dark" : "Light";

            // Load our icon resources based on the system theme.
            Icon_HyperV_Critical = GetIconResource($"HyperV_Critical_{theme}");
            Icon_HyperV_Off = GetIconResource($"HyperV_Off_{theme}");
            Icon_HyperV_Running = GetIconResource($"HyperV_Running_{theme}");
        }

        #endregion

        #region Public Static
        /// <summary>
        /// Tells <see cref="ResourceHelper"/> to try and create <see cref="ResourceManager"/> instances for the Hyper-V Tools resources.
        /// </summary>
        /// <param name="hyperVInstallFolder">The path Hyper-V Tools are installed to.</param>
        public static void LoadExternalResources(string hyperVInstallFolder)
        {
            // Check that the supplied Hyper-V Tools install folder exists.
            if (Directory.Exists(hyperVInstallFolder))
            {
                // Handle the AssemblyResolve event to manually load missing assemblies, which the Hyper-V tools might request.
                AppDomain.CurrentDomain.AssemblyResolve += (_, args) =>
                {
                    // Get a filename for the assembly, and combine it with the Hyper-V tools installs path.
                    var assemblyName = new AssemblyName(args.Name);
                    var assemblyPath = Path.Combine(hyperVInstallFolder, $"{assemblyName.Name}.dll");

                    // If the file exists on disk, load it.
                    if (File.Exists(assemblyPath))
                    {
                        try
                        {
                            return Assembly.LoadFile(assemblyPath);
                        }
                        catch
                        {
                            // We couldn't load the assembly, so we'll just swallow the exception and fall through to return null.
                        }
                    }

                    // Return null if the assembly cannot be resolved.
                    return null; 
                };

                // Attempt to load the Hyper-V assemblies we'll use for resources and create a `ResourceManager` instance for each so we
                // can load the resources from it.

                // Microsoft.Virtualization.Client.dll
                var clientAssemblyPath = Path.Combine(hyperVInstallFolder, "Microsoft.Virtualization.Client.dll");
                if (File.Exists(clientAssemblyPath))
                {
                    try
                    {
                        var clientAssembly = Assembly.LoadFile(clientAssemblyPath);
                        _clientResourceManager = new ResourceManager(@"Microsoft.Virtualization.Client.Resources.CommonResources", clientAssembly);
                    }
                    catch
                    {
                        // If we get an exception there isn't much we can do, we'll just skip loading the resource, and the app will use
                        // our fallback values.
                    }
                }

                // Microsoft.Virtualization.Client.VMBrowser.dll
                var vmBrowserAssemblyPath = Path.Combine(hyperVInstallFolder, "Microsoft.Virtualization.Client.VMBrowser.dll");
                if (File.Exists(vmBrowserAssemblyPath))
                {
                    try
                    {
                        var vmBrowserAssembly = Assembly.LoadFile(vmBrowserAssemblyPath);
                        _vmBrowserResourceManager = new ResourceManager(@"Microsoft.Virtualization.Client.VMBrowser.Resources", vmBrowserAssembly);
                    }
                    catch
                    {
                        // If we get an exception there isn't much we can do, we'll just skip loading the resource, and the app will use
                        // our fallback values.
                    }
                }
            }
        }

        #endregion

        #endregion
    }
}