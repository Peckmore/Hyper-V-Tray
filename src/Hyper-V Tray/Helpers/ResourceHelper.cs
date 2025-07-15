using HyperVTray.Resources;
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

        private static ResourceManager? _vmBrowserResourceManager;

        #endregion

        #region Construction

        static ResourceHelper()
        {
            // Load the appropriate icon, depending on whether we are running on Windows 8 and later.
            if (Environment.OSVersion.Version < new Version(8, 0))
            {
                // Before Windows 8, use the classic icon.
                Icon_HyperV = GetIconResource(Icons.Icon_Classic);
            }
            else
            {
                // Windows 8 or later, use the modern icon.
                Icon_HyperV = GetIconResource(Icons.Icon_Modern);
            }
        }

        #endregion

        #region Properties

        internal static string Command_Connect => GetStringResource("VMOpen_Name", StringsFallback.Command_Connect);
        internal static string Command_HyperVManager => GetStringResource("SnapInNode_DisplayName", StringsFallback.Menu_HyperVManager);
        internal static string Command_Pause => GetStringResource("VMPause_Name", StringsFallback.Command_Pause);
        internal static string Command_Reset => GetStringResource("VMReset_Name", StringsFallback.Command_Reset);
        internal static string Command_Resume => GetStringResource("VMResume_Name", StringsFallback.Command_Resume);
        internal static string Command_Save => GetStringResource("VMSaveState_Name", StringsFallback.Command_Save);
        internal static string Command_ShutDown => GetStringResource("VMShutDown_Name", StringsFallback.Command_ShutDown);
        internal static string Command_Start => GetStringResource("VMStart_Name", StringsFallback.Command_Start);
        internal static string Command_TurnOff => GetStringResource("VMTurnOff_Name", StringsFallback.Command_TurnOff);
        internal static Icon Icon_HyperV { get; }
        internal static string Menu_AllVirtualMachines => Strings.Menu_AllVirtualMachines;
        internal static string Message_PauseVMFailed => GetStringResource("Message_PauseVMFailed", StringsFallback.Message_PauseVMFailed);
        internal static string Message_PowerOffVMFailed => GetStringResource("Message_PowerOffVMFailed", StringsFallback.Message_PowerOffVMFailed);
        internal static string Message_ResetVMFailed => GetStringResource("Message_ResetVMFailed", StringsFallback.Message_ResetVMFailed);
        internal static string Message_ResumeVMFailed => GetStringResource("Message_ResumeVMFailed", StringsFallback.Message_ResumeVMFailed);
        internal static string Message_SaveStateVMFailed => GetStringResource("Message_SaveStateVMFailed", StringsFallback.Message_SaveStateVMFailed);
        internal static string Message_ShutDownVMFailed => GetStringResource("Message_ShutDownVMFailed", StringsFallback.Message_ShutDownVMFailed);
        internal static string Message_StartVMFailed => GetStringResource("Message_StartVMFailed", StringsFallback.Message_StartVMFailed);
        internal static string State_Critical => GetStringResource("cccccccc", StringsFallback.State_Critical);
        internal static string State_Paused => GetStringResource("DDDDDD", StringsFallback.State_Paused);
        internal static string State_Running => GetStringResource("rrrrrrr", StringsFallback.State_Running);
        internal static string State_Saved => GetStringResource("dddddddd", StringsFallback.State_Saved);
        internal static string Toast_CriticalState => StringsFallback.Toast_CriticalState;
        internal static string String_UnknownVirtualMachine => Strings.String_UnknownVirtualMachine;

        #endregion

        #region Methods

        #region Private Static

        private static Icon GetIconResource(byte[] bytes)
        {
            // Grab an icon resource from our embedded resources, and return it as an `Icon`.

            using (var ms = new MemoryStream(bytes))
            {
                return new Icon(ms);
            }
        }
        private static string GetStringResource(string resourceName, string fallbackValue)
        {
            // Attempt to get a string from the Hyper-V resources, and return the fallback value if the resource could not be found.
            return _vmBrowserResourceManager?.GetString(resourceName) ?? fallbackValue;
        }

        #endregion

        #region Internal Static

        internal static void Initialize(string hyperVInstallFolder)
        {
            // Handle the AssemblyResolve event to manually load missing assemblies, which the Hyper-V tools might request.
            AppDomain.CurrentDomain.AssemblyResolve += (sender, args) =>
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
                        // We couldn't load the assembly, so we'll just swallow the exception and return fall through to return null.
                    }
                }

                // Return null if the assembly cannot be resolved.
                return null; 
            };

            // Attempt to load the main Hyper-V assembly we'll use for resources.
            var vmBrowserAssembly = Assembly.LoadFile(Path.Combine(hyperVInstallFolder, "Microsoft.Virtualization.Client.VMBrowser.dll"));

            // Create a `ResourceManager` instance for the Hyper-V assembly so we can load the resources from it.
            _vmBrowserResourceManager = new ResourceManager(@"Microsoft.Virtualization.Client.VMBrowser.Resources", vmBrowserAssembly);
        }

        #endregion

        #endregion
    }
}