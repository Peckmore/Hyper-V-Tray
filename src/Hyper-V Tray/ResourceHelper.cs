using HyperVTray.Resources;
using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Resources;

namespace HyperVTray
{
    internal static class ResourceHelper
    {
        #region Fields

        private static ResourceManager? _vmBrowserResourceManager;

        #endregion

        #region Construction

        static ResourceHelper()
        {
            Icon_HyperV = GetIconResource(Icons.TrayIcon_Modern);
        }

        #endregion

        #region Properties

        internal static string Command_Connect => GetStringResource("VMOpen_Name", StringsFallback.Command_Connect);
        internal static string Command_Pause => GetStringResource("VMPause_Name", StringsFallback.Command_Pause);
        internal static string Command_Reset => GetStringResource("VMReset_Name", StringsFallback.Command_Reset);
        internal static string Command_Resume => GetStringResource("VMResume_Name", StringsFallback.Command_Resume);
        internal static string Command_Save => GetStringResource("VMSaveState_Name", StringsFallback.Command_Save);
        internal static string Command_ShutDown => GetStringResource("VMShutDown_Name", StringsFallback.Command_ShutDown);
        internal static string Command_Start => GetStringResource("VMStart_Name", StringsFallback.Command_Start);
        internal static string Command_TurnOff => GetStringResource("VMTurnOff_Name", StringsFallback.Command_TurnOff);
        internal static Icon Icon_HyperV { get; }
        internal static string Menu_AllVirtualMachines => GetStringResource("Menu_AllVirtualMachines", Strings.Menu_AllVirtualMachines);
        internal static string Menu_HyperVManager => GetStringResource("SnapInNode_DisplayName", StringsFallback.Menu_HyperVManager);
        internal static string Message_PauseVMFailed => GetStringResource("Message_PauseVMFailed", StringsFallback.Message_PauseVMFailed);
        internal static string Message_PowerOffVMFailed => GetStringResource("Message_PowerOffVMFailed", StringsFallback.Message_PowerOffVMFailed);
        internal static string Message_ResetVMFailed => GetStringResource("Message_ResetVMFailed", StringsFallback.Message_ResetVMFailed);
        internal static string Message_ResumeVMFailed => GetStringResource("Message_ResumeVMFailed", StringsFallback.Message_ResumeVMFailed);
        internal static string Message_SaveStateVMFailed => GetStringResource("Message_SaveStateVMFailed", StringsFallback.Message_SaveStateVMFailed);
        internal static string Message_ShutDownVMFailed => GetStringResource("Message_ShutDownVMFailed", StringsFallback.Message_ShutDownVMFailed);
        internal static string Message_StartVMFailed => GetStringResource("Message_StartVMFailed", StringsFallback.Message_StartVMFailed);
        internal static string State_Critical => GetStringResource("cccccccc", StringsFallback.State_Critical);
        internal static string State_Off => GetStringResource("fffffff", StringsFallback.State_Off);
        internal static string State_Paused => GetStringResource("DDDDDD", StringsFallback.State_Paused);
        internal static string State_Running => GetStringResource("rrrrrrr", StringsFallback.State_Running);
        internal static string State_Saved => GetStringResource("dddddddd", StringsFallback.State_Saved);
        internal static string Toast_CriticalState => StringsFallback.Toast_CriticalState;
        internal static string String_UnknownVirtualMachine => StringsFallback.String_UnknownVirtualMachine;

        #endregion

        #region Methods

        #region Private Static

        private static Icon GetIconResource(byte[] bytes)
        {
            using (var ms = new MemoryStream(bytes))
            {
                return new Icon(ms);
            }
        }
        private static string GetStringResource(string resourceName, string fallbackValue)
        {
            return _vmBrowserResourceManager?.GetString(resourceName) ?? fallbackValue;
        }

        #endregion

        #region Internal Static

        internal static void Initialize(string hyperVInstallFolder)
        {
            // Handle the AssemblyResolve event to manually load missing assemblies
            AppDomain.CurrentDomain.AssemblyResolve += (sender, args) =>
            {
                var assemblyName = new AssemblyName(args.Name);
                var assemblyPath = Path.Combine(hyperVInstallFolder, $"{assemblyName.Name}.dll");

                if (File.Exists(assemblyPath))
                {
                    return Assembly.LoadFile(assemblyPath);
                }

                return null; // Return null if the assembly cannot be resolved
            };

            var vmBrowserAssembly = Assembly.LoadFile(Path.Combine(hyperVInstallFolder, "Microsoft.Virtualization.Client.VMBrowser.dll"));
            _vmBrowserResourceManager = new ResourceManager(@"Microsoft.Virtualization.Client.VMBrowser.Resources", vmBrowserAssembly);
        }

        #endregion

        #endregion
    }
}