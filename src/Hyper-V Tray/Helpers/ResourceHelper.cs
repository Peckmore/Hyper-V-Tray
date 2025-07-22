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

        private static ResourceManager? _clientResourceManager;
        private static ResourceManager? _vmBrowserResourceManager;

        #endregion

        #region Construction

        static ResourceHelper()
        {
            // Load the appropriate icon, depending on whether we are running on Windows 8 and later.
            if (Environment.OSVersion.Version < new Version(8, 0))
            {
                // Before Windows 8, use the classic icon.
                Icon_HyperV = GetIconResource("HyperV_Classic");
            }
            else
            {
                // Windows 8 or later, use the modern icon.
                Icon_HyperV = GetIconResource("HyperV");
            }
        }

        #endregion

        #region Properties

        internal static string Button_DontReset => GetStringResource(_clientResourceManager, "ConfirmationResetButton2", StringsFallback.Button_DontReset);
        internal static string Button_DontShutDown => GetStringResource(_clientResourceManager, "ConfirmationShutdownButton2", StringsFallback.Button_DontShutDown);
        internal static string Button_DontTurnOff => GetStringResource(_clientResourceManager, "ConfirmationTurnoffButton2", StringsFallback.Button_DontTurnOff);
        internal static string Button_Reset => GetStringResource(_clientResourceManager, "ConfirmationResetButton1", StringsFallback.Button_Reset);
        internal static string Button_ShutDown => GetStringResource(_clientResourceManager, "ConfirmationShutdownButton1", StringsFallback.Button_ShutDown);
        internal static string Button_TurnOff => GetStringResource(_clientResourceManager, "ConfirmationTurnoffButton1", StringsFallback.Button_TurnOff);
        internal static string Command_Connect => GetStringResource(_vmBrowserResourceManager, "VMOpen_Name", StringsFallback.Command_Connect);
        internal static string Command_HyperVManager => GetStringResource(_vmBrowserResourceManager, "SnapInNode_DisplayName", StringsFallback.Menu_HyperVManager);
        internal static string Command_Pause => GetStringResource(_vmBrowserResourceManager, "VMPause_Name", StringsFallback.Command_Pause);
        internal static string Command_Reset => GetStringResource(_vmBrowserResourceManager, "VMReset_Name", StringsFallback.Command_Reset);
        internal static string Command_Resume => GetStringResource(_vmBrowserResourceManager, "VMResume_Name", StringsFallback.Command_Resume);
        internal static string Command_Save => GetStringResource(_vmBrowserResourceManager, "VMSaveState_Name", StringsFallback.Command_Save);
        internal static string Command_ShutDown => GetStringResource(_vmBrowserResourceManager, "VMShutDown_Name", StringsFallback.Command_ShutDown);
        internal static string Command_Start => GetStringResource(_vmBrowserResourceManager, "VMStart_Name", StringsFallback.Command_Start);
        internal static string Command_TurnOff => GetStringResource(_vmBrowserResourceManager, "VMTurnOff_Name", StringsFallback.Command_TurnOff);
        internal static Icon Icon_HyperV { get; }
        internal static string Menu_AllVirtualMachines => Strings.Menu_AllVirtualMachines;
        internal static string Message_ConfirmationReset => GetStringResource(_clientResourceManager, "ConfirmationReset", StringsFallback.Message_ConfirmationReset);
        internal static string Message_ConfirmationResetMultiple => GetStringResource(_clientResourceManager, "ConfirmationResetMultiple", StringsFallback.Message_ConfirmationResetMultiple);
        internal static string Message_ConfirmationShutDown => GetStringResource(_clientResourceManager, "ConfirmationShutdown", StringsFallback.Message_ConfirmationShutDown);
        internal static string Message_ConfirmationShutDownMultiple => GetStringResource(_clientResourceManager, "ConfirmationShutdownMultiple", StringsFallback.Message_ConfirmationShutDownMultiple);
        internal static string Message_ConfirmationTurnOff => GetStringResource(_clientResourceManager, "ConfirmationTurnoff", StringsFallback.Message_ConfirmationTurnOff);
        internal static string Message_ConfirmationTurnOffMultiple => GetStringResource(_clientResourceManager, "ConfirmationTurnoffMultiple", StringsFallback.Message_ConfirmationTurnOffMultiple);
        internal static string Message_OpenVMConnectFailed => GetStringResource(_vmBrowserResourceManager, "Message_OpenVMConnectFailed", StringsFallback.Message_OpenVMConnectFailed);
        internal static string Message_PauseVMFailed => GetStringResource(_vmBrowserResourceManager, "Message_PauseVMFailed_Format", StringsFallback.Message_PauseVMFailed);
        internal static string Message_PauseVMFailedMultiple => GetStringResource(_vmBrowserResourceManager, "Message_PauseVMFailed", StringsFallback.Message_PauseVMFailedMultiple);
        internal static string Message_PowerOffVMFailed => GetStringResource(_vmBrowserResourceManager, "Message_PowerOffFailed_Format", StringsFallback.Message_PowerOffVMFailed);
        internal static string Message_PowerOffVMFailedMultiple => GetStringResource(_vmBrowserResourceManager, "Message_PowerOffVMFailed", StringsFallback.Message_PowerOffVMFailed);
        internal static string Message_ResetVMFailed => GetStringResource(_vmBrowserResourceManager, "Message_ResetVMFailed_Format", StringsFallback.Message_ResetVMFailed);
        internal static string Message_ResetVMFailedMultiple => GetStringResource(_vmBrowserResourceManager, "Message_ResetVMFailed", StringsFallback.Message_ResetVMFailedMultiple);
        internal static string Message_ResumeVMFailed => GetStringResource(_vmBrowserResourceManager, "Message_ResumeVMFailed", StringsFallback.Message_ResumeVMFailedMultiple);
        internal static string Message_ResumeVMFailedMultiple => GetStringResource(_vmBrowserResourceManager, "Message_ResumeVMFailed", StringsFallback.Message_ResumeVMFailedMultiple);
        internal static string Message_SaveStateVMFailed => GetStringResource(_vmBrowserResourceManager, "Message_SaveStateVMFailed_Format", StringsFallback.Message_SaveStateVMFailed);
        internal static string Message_SaveStateVMFailedMultiple => GetStringResource(_vmBrowserResourceManager, "Message_SaveStateVMFailed", StringsFallback.Message_SaveStateVMFailedMultiple);
        internal static string Message_ShutDownVMFailed => GetStringResource(_vmBrowserResourceManager, "Message_ShutDownVMFailed", StringsFallback.Message_ShutDownVMFailedMultiple);
        internal static string Message_ShutDownVMFailedMultiple => GetStringResource(_vmBrowserResourceManager, "Message_ShutDownVMFailed", StringsFallback.Message_ShutDownVMFailedMultiple);
        internal static string Message_StartVMFailed => GetStringResource(_vmBrowserResourceManager, "Message_StartVMFailed_Format", StringsFallback.Message_StartVMFailed);
        internal static string Message_StartVMFailedMultiple => GetStringResource(_vmBrowserResourceManager, "Message_StartVMFailed", StringsFallback.Message_StartVMFailedMultiple);
        internal static string State_Critical => GetStringResource(_vmBrowserResourceManager, "VMOnFailureStateFormat", StringsFallback.State_Critical);
        internal static string State_Off => GetStringResource(_vmBrowserResourceManager, "DDDDDD", StringsFallback.State_Off);
        internal static string State_Paused => GetStringResource(_vmBrowserResourceManager, "DDDDDD", StringsFallback.State_Paused);
        internal static string State_Running => GetStringResource(_vmBrowserResourceManager, "rrrrrrr", StringsFallback.State_Running);
        internal static string State_Saved => GetStringResource(_vmBrowserResourceManager, "dddddddd", StringsFallback.State_Saved);
        internal static string State_Unknown => GetStringResource(_vmBrowserResourceManager, "dddddddd", StringsFallback.State_Unknown);
        internal static string Title_ResetMachine => GetStringResource(_clientResourceManager, "ConfirmationResetTitle", StringsFallback.Title_ResetMachine);
        internal static string Title_ShutDownMachine => GetStringResource(_clientResourceManager, "ConfirmationShutdownTitle", StringsFallback.Title_ShutDownMachine);
        internal static string Title_TurnOffMachine => GetStringResource(_clientResourceManager, "ConfirmationTurnoffTitle", StringsFallback.Title_TurnOffMachine);
        internal static string Toast_CriticalState => Strings.Toast_CriticalState;
        internal static string String_UnknownVirtualMachine => Strings.String_UnknownVirtualMachine;

        #endregion

        #region Methods

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

        #endregion

        #region Internal Static

        internal static void Initialize(string hyperVInstallFolder)
        {
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
                            // We couldn't load the assembly, so we'll just swallow the exception and return fall through to return null.
                        }
                    }

                    // Return null if the assembly cannot be resolved.
                    return null; 
                };

                // Attempt to load the Hyper-V assemblies we'll use for resources and create a `ResourceManager` instance for each so we
                // can load the resources from it.
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