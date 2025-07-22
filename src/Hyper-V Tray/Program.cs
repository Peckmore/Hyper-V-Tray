using HyperVTray.Helpers;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace HyperVTray
{
    internal static class Program
    {
        #region Constants

        #region Private

        private const int BalloonTipTimeout = 2500;

        #endregion

        #region Public

        public const string ApplicationName = @"Hyper-V Tray";

        #endregion

        #endregion

        #region Fields

        private static string? _hyperVInstallFolder;
        private static string? _vmConnectPath;
        private static string? _vmManagerPath;
        private static readonly ContextMenu ContextMenu;
        private static readonly NotifyIcon NotifyIcon;

        #endregion

        #region Construction

        static Program()
        {
            // Create our notification icon to place in the tray.
            NotifyIcon = new NotifyIcon();
            NotifyIcon.DoubleClick += NotifyIcon_DoubleClick;
            NotifyIcon.MouseClick += NotifyIcon_MouseClick;
            
            // Create our context menu instance which we'll display whenever the icon is clicked.
            ContextMenu = new ContextMenu();

            // Listen for display settings changes so we can handle DPI changes.
            SystemEvents.DisplaySettingsChanged += SystemEvents_DisplaySettingsChanged;
        }

        #endregion

        #region Methods

        #region Event Handlers

        private static void HyperVHelper_VirtualMachineStateChanged(object? sender, VirtualMachineStateChangedEventArgs e)
        {
            ShowToast(e.Name, e.State, e.Critical);
        }
        private static void NotifyIcon_DoubleClick(object? sender, EventArgs e)
        {
            OpenHyperVManager();
        }
        private static void NotifyIcon_MouseClick(object? sender, MouseEventArgs e)
        {
            GenerateContextMenu();
            NotifyIcon.ShowContextMenu(ContextMenu);
        }
        private static void SystemEvents_DisplaySettingsChanged(object? sender, EventArgs e)
        {
            Debug.WriteLine("Display settings change detected.");

            SetTrayIcon();
        }

        #endregion

        #region Private Static

        private static bool ConfirmAction(VmState state, bool multiple)
        {
            var confirm = true;

            if (state is VmState.Disabled or VmState.Reset or VmState.ShutDown)
            {
                string? button1Text = null;
                string? button2Text = null;
                string? heading = null;
                string? message = null;

                switch (state)
                {
                    case VmState.Disabled:
                        button1Text = ResourceHelper.Button_TurnOff;
                        button2Text = ResourceHelper.Button_DontTurnOff;
                        heading = ResourceHelper.Title_TurnOffMachine;
                        message = multiple ? ResourceHelper.Message_ConfirmationTurnOffMultiple : ResourceHelper.Message_ConfirmationTurnOff;
                        break;
                    case VmState.Reset:
                        button1Text = ResourceHelper.Button_Reset;
                        button2Text = ResourceHelper.Button_DontReset;
                        heading = ResourceHelper.Title_ResetMachine;
                        message = multiple ? ResourceHelper.Message_ConfirmationResetMultiple : ResourceHelper.Message_ConfirmationReset;
                        break;
                    case VmState.ShutDown:
                        button1Text = ResourceHelper.Button_ShutDown;
                        button2Text = ResourceHelper.Button_DontShutDown;
                        heading = ResourceHelper.Title_ShutDownMachine;
                        message = multiple ? ResourceHelper.Message_ConfirmationShutDownMultiple : ResourceHelper.Message_ConfirmationShutDown;
                        break;
                }

                var button1 = new TaskDialogButton(button1Text);
                var button2 = new TaskDialogButton(button2Text);

                var taskDialogPage = new TaskDialogPage
                {
                    AllowCancel = true,
                    AllowMinimize = false,
                    Buttons = new() { button1, button2 },
                    Caption = heading,
                    DefaultButton = button2,
                    Heading = message,
                    Icon = TaskDialogIcon.Warning
                };
                confirm = TaskDialog.ShowDialog(taskDialogPage) == button1;
            }

            return confirm;
        }
        private static void ControlAllVirtualMachines(VmState state)
        {
            var hasErrors = false;

            if (ConfirmAction(state, true))
            {
                foreach (var virtualMachine in HyperVHelper.GetVirtualMachines())
                {
                    // Get the VM name.
                    var virtualMachineName = virtualMachine["ElementName"].ToString();

                    // Set the VM state.
                    if (!ControlVirtualMachine(virtualMachineName, state, false, false))
                    {
                        hasErrors = true;
                    }
                }
            }

            if (hasErrors)
            {
                var errorMessage = state switch
                {
                    VmState.Disabled => ResourceHelper.Message_PowerOffVMFailedMultiple,
                    VmState.Enabled => ResourceHelper.Message_StartVMFailedMultiple,
                    VmState.Offline => ResourceHelper.Message_SaveStateVMFailedMultiple,
                    VmState.Quiesce => ResourceHelper.Message_PauseVMFailedMultiple,
                    VmState.Reset => ResourceHelper.Message_ResetVMFailedMultiple,
                    VmState.Resuming => ResourceHelper.Message_ResumeVMFailedMultiple,
                    VmState.ShutDown => ResourceHelper.Message_ShutDownVMFailedMultiple,
                    _ => string.Empty
                };

                ShowError(errorMessage);
            }
        }
        private static bool ControlVirtualMachine(string virtualMachineName, VmState state, bool showErrorMessage = true, bool promptToConfirm = true)
        {
            var result = false;

            if (!promptToConfirm || ConfirmAction(state, false))
            {
                result = state switch
                {
                    VmState.Disabled or
                    VmState.Enabled or
                    VmState.Offline or
                    VmState.Quiesce or
                    VmState.Reset => HyperVHelper.RequestVmStateChange(virtualMachineName, state),
                    VmState.Resuming => HyperVHelper.RequestVmStateChange(virtualMachineName, VmState.Enabled),
                    VmState.ShutDown => HyperVHelper.RequestVmShutdown(virtualMachineName),
                    _ => false
                };

                if (!result && showErrorMessage)
                {
                    var errorMessage = state switch
                    {
                        VmState.Disabled => ResourceHelper.Message_PowerOffVMFailed,
                        VmState.Enabled => ResourceHelper.Message_StartVMFailed,
                        VmState.Offline => ResourceHelper.Message_SaveStateVMFailed,
                        VmState.Quiesce => ResourceHelper.Message_PauseVMFailed,
                        VmState.Reset => ResourceHelper.Message_ResetVMFailed,
                        VmState.Resuming => ResourceHelper.Message_ResumeVMFailed,
                        VmState.ShutDown => ResourceHelper.Message_ShutDownVMFailed,
                        _ => string.Empty
                    };

                    ShowError(string.Format(errorMessage, $"'{virtualMachineName}'"));
                }
            }

            return result;
        }
        private static void ConnectToVirtualMachine(string virtualMachineName)
        {
            try
            {
            var processInfo = new ProcessStartInfo
            {
                FileName = _vmConnectPath!,
                Arguments = $"localhost \"{virtualMachineName}\"",
            };
            Process.Start(processInfo);
        }
            catch
            {
                ShowError(ResourceHelper.Message_OpenVMConnectFailed);
            }
        }
        private static void ExitApplication()
        {
            NotifyIcon.Visible = false;
            Application.Exit();
        }
        private static void GenerateContextMenu()
        {
            // Clear the context menu.
            ContextMenu.MenuItems.Clear();

            // Get all VMs.
            var virtualMachines = HyperVHelper.GetVirtualMachines();

            // We'll set some flags as we iterate through the virtual machines to determine if any are in certain states, then use these
            // flags to build our sub-menu for commands to run against all relevant virtual machines.
            var isAnyOffOrSaved = false;
            var isAnyPaused = false;
            var isAnyRunning = false;

            // Create a menu entry for each VM.
            foreach (var virtualMachine in virtualMachines)
            {
                // Get the VM name.
                var virtualMachineName = virtualMachine["ElementName"].ToString();

                if (virtualMachineName != null)
                {
                    // Get the VM state.
                    var virtualMachineStatus = (VmState)Convert.ToInt32(virtualMachine["EnabledState"]);

                    // Generate the menu entry title for the VM.
                    var virtualMachineMenuTitle = virtualMachineName;
                    if (virtualMachineStatus != VmState.Disabled) // Stopped
                    {
                        virtualMachineMenuTitle += $" [{HyperVHelper.VmStateToString(virtualMachineStatus)}]";
                    }

                    // Create VM menu item.
                    var virtualMachineMenu = new MenuItem(virtualMachineMenuTitle) { Name = virtualMachineName };
                    var isPaused = HyperVHelper.IsVmPaused(virtualMachineStatus);
                    var isOffOrSaved = HyperVHelper.IsVmOff(virtualMachineStatus) || HyperVHelper.IsVmSaved(virtualMachineStatus);

                    // Now generate control menu items for VM.

                    // Connect
                    if (_vmConnectPath != null)
                    {
                        var connectMenuItem = new MenuItem(ResourceHelper.Command_Connect, (_, _) => ConnectToVirtualMachine(virtualMachineName));
                        virtualMachineMenu.MenuItems.Add(connectMenuItem);
                        virtualMachineMenu.MenuItems.Add(new MenuItem("-"));
                    }

                    if (isOffOrSaved)
                    {
                        // We have at least one machine that is powered off, so we can set this flag to true.
                        isAnyOffOrSaved = true;

                        // Start
                        var startMenuItem = new MenuItem(ResourceHelper.Command_Start, (_, _) => ControlVirtualMachine(virtualMachineName, VmState.Enabled));
                        virtualMachineMenu.MenuItems.Add(startMenuItem);
                    }
                    else
                    {
                        // Turn Off
                        var stopMenuItem = new MenuItem(ResourceHelper.Command_TurnOff, (_, _) => ControlVirtualMachine(virtualMachineName, VmState.Disabled));
                        virtualMachineMenu.MenuItems.Add(stopMenuItem);

                        // Shut Down
                        if (!isPaused)
                        {
                            // We have at least one machine that is running, so we can set this flag to true.
                            isAnyRunning = true;

                            var shutMenuDownItem = new MenuItem(ResourceHelper.Command_ShutDown, (_, _) => ControlVirtualMachine(virtualMachineName, VmState.ShutDown));
                            virtualMachineMenu.MenuItems.Add(shutMenuDownItem);
                        }

                        // Save
                        var saveMenuStateItem = new MenuItem(ResourceHelper.Command_Save, (_, _) => ControlVirtualMachine(virtualMachineName, VmState.Offline));
                        virtualMachineMenu.MenuItems.Add(saveMenuStateItem);

                        virtualMachineMenu.MenuItems.Add(new MenuItem("-"));
                        if (isPaused)
                        {
                            // We have at least one machine that is paused, so we can set this flag to true.
                            isAnyPaused = true;

                            // Resume
                            var resumeMenuItem = new MenuItem(ResourceHelper.Command_Resume, (_, _) => ControlVirtualMachine(virtualMachineName, VmState.Resuming));
                            virtualMachineMenu.MenuItems.Add(resumeMenuItem);
                        }
                        else
                        {
                            // Pause
                            var pauseMenuItem = new MenuItem(ResourceHelper.Command_Pause, (_, _) => ControlVirtualMachine(virtualMachineName, VmState.Quiesce));
                            virtualMachineMenu.MenuItems.Add(pauseMenuItem);
                        }

                        // Reset
                        var resetMenuItem = new MenuItem(ResourceHelper.Command_Reset, (_, _) => ControlVirtualMachine(virtualMachineName, VmState.Reset));
                        virtualMachineMenu.MenuItems.Add(resetMenuItem);
                    }

                    // Add VM menu item to root menu.
                    ContextMenu.MenuItems.Add(virtualMachineMenu);
                }
            }

            if (virtualMachines.Any())
            {
                ContextMenu.MenuItems.Add(new MenuItem("-"));

                // Create a root menu item for the `All Virtual Machines` entry.
                var vmItem = new MenuItem(ResourceHelper.Menu_AllVirtualMachines);

                var subItems = new List<MenuItem>();

                // Start
                if (isAnyOffOrSaved)
                {
                    var startAllMenuItem = new MenuItem(ResourceHelper.Command_Start, (_, _) => ControlAllVirtualMachines(VmState.Enabled));
                    subItems.Add(startAllMenuItem);
                }

                // Turn Off
                if (isAnyRunning || isAnyPaused)
                {
                    var stopAllMenuItem = new MenuItem(ResourceHelper.Command_TurnOff, (_, _) => ControlAllVirtualMachines(VmState.Disabled));
                    subItems.Add(stopAllMenuItem);
                }

                // Shut Down
                if (isAnyRunning)
                {
                    var shutDownAllMenuItem = new MenuItem(ResourceHelper.Command_ShutDown, (_, _) => ControlAllVirtualMachines(VmState.ShutDown));
                    subItems.Add(shutDownAllMenuItem);
                }

                // Save
                if (isAnyRunning || isAnyPaused)
                {
                    var saveAllMenuItem = new MenuItem(ResourceHelper.Command_Save, (_, _) => ControlAllVirtualMachines(VmState.Offline));
                    subItems.Add(saveAllMenuItem);
                }

                if (subItems.Any() && isAnyRunning || isAnyPaused)
                {
                    subItems.Add(new MenuItem("-"));
                }

                // Pause
                if (isAnyRunning)
                {
                    var pauseAllMenuItem = new MenuItem(ResourceHelper.Command_Pause, (_, _) => ControlAllVirtualMachines(VmState.Quiesce));
                    subItems.Add(pauseAllMenuItem);
                }

                // Resume
                if (isAnyPaused)
                {
                    var resumeAllMenuItem = new MenuItem(ResourceHelper.Command_Resume, (_, _) => ControlAllVirtualMachines(VmState.Resuming));
                    subItems.Add(resumeAllMenuItem);
                }

                // Reset
                if (isAnyRunning || isAnyPaused)
                {
                    var resetAllMenuItem = new MenuItem(ResourceHelper.Command_Reset, (_, _) => ControlAllVirtualMachines(VmState.Reset));
                    subItems.Add(resetAllMenuItem);
                }

                vmItem.MenuItems.AddRange(subItems.ToArray());

                // Add the VM to the context menu.
                ContextMenu.MenuItems.Add(vmItem);
            }

            // Add `Hyper-V Manager` menu item.
            if (_vmManagerPath != null)
            {
                ContextMenu.MenuItems.Add(new MenuItem("-"));
                var managerItem = new MenuItem(ResourceHelper.Command_HyperVManager);
                managerItem.Click += (_, _) => OpenHyperVManager();
                ContextMenu.MenuItems.Add(managerItem);
            }

            // Add `Exit` menu item.
            ContextMenu.MenuItems.Add(new MenuItem("-"));
            var exitItem = new MenuItem("Exit");
            exitItem.Click += (_, _) => ExitApplication();
            ContextMenu.MenuItems.Add(exitItem);
        }
        [STAThread]
        private static void Main()
        {
            // Application setup.
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // Show our tray icon.
            NotifyIcon.Visible = true;
            SetTrayIcon();

            // Detect Hyper-V components.
            _hyperVInstallFolder = @$"{Environment.GetEnvironmentVariable("ProgramFiles")}\Hyper-V\";
            if (Directory.Exists(_hyperVInstallFolder))
            {
                _vmConnectPath = $@"{Environment.GetEnvironmentVariable("SYSTEMROOT")}\System32\vmconnect.exe";
                if (!File.Exists(_vmConnectPath))
                {
                    _vmConnectPath = null;
                }

                _vmManagerPath = $@"{Environment.GetEnvironmentVariable("SYSTEMROOT")}\System32\virtmgmt.msc";
                if (!File.Exists(_vmManagerPath))
                {
                    _vmManagerPath = null;
                }
            }

            // Initialize our helpers.
            HyperVHelper.Initialize();
            HyperVHelper.VirtualMachineStateChanged += HyperVHelper_VirtualMachineStateChanged;
            ResourceHelper.Initialize(_hyperVInstallFolder);

            // Run the application.
            Application.Run();
        }
        private static void OpenHyperVManager()
        {
            var processInfo = new ProcessStartInfo
            {
                FileName = @$"{Environment.GetEnvironmentVariable("SYSTEMROOT")}\System32\mmc.exe",
                Arguments = _vmManagerPath!,
                WorkingDirectory = _hyperVInstallFolder,
                UseShellExecute = true,
                Verb = @"runas"
            };
            Process.Start(processInfo);
        }
        private static void SetTrayIcon()
        {
            NotifyIcon.Icon = new Icon(ResourceHelper.Icon_HyperV, SystemInformation.SmallIconSize);
            Debug.WriteLine($"Icon size: {SystemInformation.SmallIconSize.Width}x{SystemInformation.SmallIconSize.Height}");
        }
        private static void ShowError(string heading, string text = "")
        {
            var taskDialogPage = new TaskDialogPage
            {
                AllowCancel = false,
                AllowMinimize = false,
                Buttons = new() { TaskDialogButton.Close },
                Caption = ApplicationName,
                Heading = heading,
                Icon = TaskDialogIcon.Error,
                Text = text
            };
            TaskDialog.ShowDialog(taskDialogPage);
        }
        private static void ShowToast(string virtualMachineName, VmState vmState, bool isCritical)
        {
            var status = HyperVHelper.VmStateToString(vmState);

            if (isCritical)
            {
                status = $"{ResourceHelper.Toast_CriticalState}\n{status}";
            }

            NotifyIcon.ShowBalloonTip(BalloonTipTimeout, virtualMachineName, status, isCritical ? ToolTipIcon.Error : ToolTipIcon.Info);
        }

        #endregion

        #endregion
    }
}