using HyperVTray.Helpers;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;

namespace HyperVTray
{
    internal static class Program
    {
        #region Constants

        #region Private

        private const int BalloonTipTimeout = 2500;

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

            // Get the Assembly Title, or application friendly name as a fallback.
            ApplicationName = Assembly.GetExecutingAssembly().GetCustomAttribute<AssemblyTitleAttribute>()?.Title ?? AppDomain.CurrentDomain.FriendlyName;
        }

        #endregion

        #region Properties

        /// <summary>
        /// The display name of the application.
        /// </summary>
        public static string ApplicationName { get; }

        #endregion

        #region Methods

        #region Event Handlers

        private static void HyperVHelper_VirtualMachineStateChanged(object? sender, VirtualMachineStateChangedEventArgs e)
        {
            // We've received an event that a VM has changed to a state that we should notify the user for, so we'll show a toast to let
            // the user know.
            ShowToast(e.Name, e.State);
        }
        private static void NotifyIcon_DoubleClick(object? sender, EventArgs e)
        {
            // The user has double-clicked the tray icon, lets open Hyper-V Manager.
            OpenHyperVManager();
        }
        private static void NotifyIcon_MouseClick(object? sender, MouseEventArgs e)
        {
            // The user has single-clicked the tray icon, so we'll generate the context menu...
            GenerateContextMenu();

            // ...and then show it.
            NotifyIcon.ShowContextMenu(ContextMenu);

            // Once the menu is gone we'll clear the menu items, as these are rebuilt every time, so no point them hanging around.
            ContextMenu.MenuItems.Clear();
        }
        private static void SystemEvents_DisplaySettingsChanged(object? sender, EventArgs e)
        {
            // Display settings have changed - in-case it's a DPI change, we'll regenerate the tray icon.
            SetTrayIcon();
        }

        #endregion

        #region Private Static

        private static bool ConfirmAction(VirtualMachineState state, bool multiple)
        {
            // We want to change the state of a VM to something that the user needs to confirm first. We'll see what state has been
            // requested, then show the appropriate prompt (if required).

            // Create a flag to indicate whether the user has confirmed. We default to true so that if the state requested doesn't
            // require a prompt, we can just return true.
            var confirmed = true;

            // Check if the state is one of the ones we need to prompt for.
            if (state is VirtualMachineState.Disabled
                      or VirtualMachineState.Reset
                      or VirtualMachineState.ShutDown)
            {
                // Create variables for what we'll display in the Task Dialog.
                string? button1Text = null;
                string? button2Text = null;
                string? heading = null;
                string? message = null;

                // Based on the state, set our Task Dialog text fields. We use the `multiple` parameter to determine whether to show the
                // single or multiple virtual machine variants of the dialog message.
                switch (state)
                {
                    case VirtualMachineState.Disabled:
                        button1Text = ResourceHelper.Button_TurnOff;
                        button2Text = ResourceHelper.Button_DontTurnOff;
                        heading = ResourceHelper.Title_TurnOffMachine;
                        message = multiple ? ResourceHelper.Message_ConfirmationTurnOffMultiple : ResourceHelper.Message_ConfirmationTurnOff;
                        break;
                    case VirtualMachineState.Reset:
                        button1Text = ResourceHelper.Button_Reset;
                        button2Text = ResourceHelper.Button_DontReset;
                        heading = ResourceHelper.Title_ResetMachine;
                        message = multiple ? ResourceHelper.Message_ConfirmationResetMultiple : ResourceHelper.Message_ConfirmationReset;
                        break;
                    case VirtualMachineState.ShutDown:
                        button1Text = ResourceHelper.Button_ShutDown;
                        button2Text = ResourceHelper.Button_DontShutDown;
                        heading = ResourceHelper.Title_ShutDownMachine;
                        message = multiple ? ResourceHelper.Message_ConfirmationShutDownMultiple : ResourceHelper.Message_ConfirmationShutDown;
                        break;
                }

                // Create the buttons we'll show on the Task Dialog.
                var buttonConfirm = new TaskDialogButton(button1Text);
                var buttonCancel = new TaskDialogButton(button2Text);

                // Now create our Task Dialog, using the message fields and buttons we determined above.
                var taskDialogPage = new TaskDialogPage
                {
                    AllowCancel = true,
                    AllowMinimize = false,
                    Buttons = new() { buttonConfirm, buttonCancel },
                    Caption = heading,
                    DefaultButton = buttonCancel,
                    Heading = message,
                    Icon = TaskDialogIcon.Warning
                };

                // Show the Task Dialog, and grab the button the user pressed.
                var result = TaskDialog.ShowDialog(taskDialogPage);

                // Update our flag to indicate whether the user pressed the "confirm" button or not.
                confirmed = result == buttonConfirm;
            }

            // Return our confirmation result.
            return confirmed;
        }
        private static void ControlAllVirtualMachines(VirtualMachineState state)
        {
            // We're going to attempt to set the state of all VMs to the requested state.

            // Create a flag to store the result of attempting to set the state.
            var hasErrors = false;

            // First we call `ConfirmAction` to show a message to the user confirming the state change request.
            if (ConfirmAction(state, true))
            {
                // The user has confirmed the state change request (or no confirmation was needed), so proceed to try and set the state.

                // Now iterate through each VM, and set the state on each one.
                foreach (var virtualMachine in HyperVHelper.GetVirtualMachines())
                {
                    // Get the VM name.
                    var virtualMachineName = virtualMachine["ElementName"].ToString();

                    // Set the VM state.
                    if (string.IsNullOrWhiteSpace(virtualMachineName) || !ControlVirtualMachine(virtualMachineName, state, false, false))
                    {
                        // Attempting to set the state has failed, so set our flag to true so once we've finished our loop we can let the
                        // user know.
                        hasErrors = true;
                    }
                }
            }

            // Check our flag to see if one or more VMs failed to change state.
            if (hasErrors)
            {
                // Based on the state requested for all VMs, determine which error message to show.
                var errorMessage = state switch
                {
                    VirtualMachineState.Disabled => ResourceHelper.Message_PowerOffVMFailedMultiple,
                    VirtualMachineState.Enabled => ResourceHelper.Message_StartVMFailedMultiple,
                    VirtualMachineState.Offline => ResourceHelper.Message_SaveStateVMFailedMultiple,
                    VirtualMachineState.Quiesce => ResourceHelper.Message_PauseVMFailedMultiple,
                    VirtualMachineState.Reset => ResourceHelper.Message_ResetVMFailedMultiple,
                    VirtualMachineState.Resuming => ResourceHelper.Message_ResumeVMFailedMultiple,
                    VirtualMachineState.ShutDown => ResourceHelper.Message_ShutDownVMFailedMultiple,
                    _ => string.Empty
                };

                // Now show the error using the chosen message.
                ShowError(errorMessage);
            }
        }
        private static bool ControlVirtualMachine(string virtualMachineName, VirtualMachineState state, bool showErrorMessage = true, bool promptToConfirm = true)
        {
            // We're going to attempt to set the state of a single VM to the requested state.

            // Create a flag to store the result of attempting to set the state.
            var result = false;

            // If `promptToConfirm` is true, we call `ConfirmAction` to show a message to the user confirming the state change request.
            // This is generally true when controlling a single VM, and false when called by "ControlAllVirtualMachines", as a single prompt
            // is shown for the entire batch task.
            if (!promptToConfirm || ConfirmAction(state, false))
            {
                // Depending on the requested state, we perform a different action.
                // - For most requested states, we call `RequestVmStateChange`, passing in the requested state.
                // - For `VmState.Resuming`, we again call `RequestVmStateChange`, but request the `Enabled` state, as `Resuming` is just
                //   used internally by us to different between Start and Resume (as they have different error messages).
                // - For `VmState.ShutDown`, we call `RequestVmShutdown`.
                result = state switch
                {
                    VirtualMachineState.Disabled or
                    VirtualMachineState.Enabled or
                    VirtualMachineState.Offline or
                    VirtualMachineState.Quiesce or
                    VirtualMachineState.Reset => HyperVHelper.RequestVirtualMachineStateChange(virtualMachineName, state),
                    VirtualMachineState.Resuming => HyperVHelper.RequestVirtualMachineStateChange(virtualMachineName, VirtualMachineState.Enabled),
                    VirtualMachineState.ShutDown => HyperVHelper.RequestVirtualMachineShutdown(virtualMachineName),
                    _ => false
                };

                // Check our flag to see if one or more VMs failed to change state. If `showErrorMessage` is true, we call `ShowError` to
                // show an error message to the user. This is generally true when controlling a single VM, and false when called by
                // "ControlAllVirtualMachines", as a single error message is shown for the entire batch task.
                if (!result && showErrorMessage)
                {
                    // Based on the state requested for the VM, determine which error message to show.
                    var errorMessage = state switch
                    {
                        VirtualMachineState.Disabled => ResourceHelper.Message_PowerOffVMFailed,
                        VirtualMachineState.Enabled => ResourceHelper.Message_StartVMFailed,
                        VirtualMachineState.Offline => ResourceHelper.Message_SaveStateVMFailed,
                        VirtualMachineState.Quiesce => ResourceHelper.Message_PauseVMFailed,
                        VirtualMachineState.Reset => ResourceHelper.Message_ResetVMFailed,
                        VirtualMachineState.Resuming => ResourceHelper.Message_ResumeVMFailed,
                        VirtualMachineState.ShutDown => ResourceHelper.Message_ShutDownVMFailed,
                        _ => string.Empty
                    };

                    // Now show the error using the chosen message.
                    ShowError(string.Format(errorMessage, $"'{virtualMachineName}'"));
                }
            }

            return result;
        }
        private static void ConnectToVirtualMachine(string virtualMachineName)
        {
            // We'll attempt to run "vmconnect.exe" to connect to, and provide control of, a virtual machine.
            try
            {
                // Set our process info, including the name of the virtual machine to connect to.
                var processInfo = new ProcessStartInfo
                {
                    FileName = _vmConnectPath!,
                    Arguments = $"localhost \"{virtualMachineName}\"",
                };

                // Start the application.
                Process.Start(processInfo);
            }
            catch
            {
                // If something happened then there isn't much we can do, so we'll just show a generic error message to let the user know
                // that we couldn't connect. At this point the user should use Hyper-V Manager to determine if there is something wrong.
                ShowError(ResourceHelper.Message_OpenVMConnectFailed);
            }
        }
        private static void ExitApplication()
        {
            // Hide our icon so it isn't left lingering when the application closes.
            NotifyIcon.Visible = false;

            // Now quit the application.
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
                    var virtualMachineState = (VirtualMachineState)Convert.ToInt32(virtualMachine["EnabledState"]);

                    // Generate the menu entry title for the VM.
                    var virtualMachineMenuTitle = virtualMachineName;
                    if (virtualMachineState != VirtualMachineState.Disabled) // Stopped
                    {
                        virtualMachineMenuTitle += $" ({HyperVHelper.VmStateToString(virtualMachineState)})";
                    }

                    // Create VM menu item.
                    var virtualMachineMenu = new MenuItem(virtualMachineMenuTitle) { Name = virtualMachineName };
                    var isPaused = HyperVHelper.IsPausedState(virtualMachineState);
                    var isOffOrSaved = HyperVHelper.IsOffState(virtualMachineState) || HyperVHelper.IsSavedState(virtualMachineState);

                    // Now generate control menu items for VM.

                    // Connect
                    if (_vmConnectPath != null)
                    {
                        // Only generate if we found the tool installed.

                        var connectMenuItem = new MenuItem(ResourceHelper.Command_Connect, (_, _) => ConnectToVirtualMachine(virtualMachineName));
                        virtualMachineMenu.MenuItems.Add(connectMenuItem);
                        virtualMachineMenu.MenuItems.Add(new MenuItem("-"));
                    }

                    if (isOffOrSaved)
                    {
                        // We have at least one machine that is powered off, so we can set this flag to true.
                        isAnyOffOrSaved = true;

                        // Start
                        var startMenuItem = new MenuItem(ResourceHelper.Command_Start, (_, _) => ControlVirtualMachine(virtualMachineName, VirtualMachineState.Enabled));
                        virtualMachineMenu.MenuItems.Add(startMenuItem);
                    }
                    else
                    {
                        // Turn Off
                        var stopMenuItem = new MenuItem(ResourceHelper.Command_TurnOff, (_, _) => ControlVirtualMachine(virtualMachineName, VirtualMachineState.Disabled));
                        virtualMachineMenu.MenuItems.Add(stopMenuItem);

                        // Shut Down
                        if (!isPaused)
                        {
                            // We have at least one machine that is running, so we can set this flag to true.
                            isAnyRunning = true;

                            var shutMenuDownItem = new MenuItem(ResourceHelper.Command_ShutDown, (_, _) => ControlVirtualMachine(virtualMachineName, VirtualMachineState.ShutDown));
                            virtualMachineMenu.MenuItems.Add(shutMenuDownItem);
                        }

                        // Save
                        var saveMenuStateItem = new MenuItem(ResourceHelper.Command_Save, (_, _) => ControlVirtualMachine(virtualMachineName, VirtualMachineState.Offline));
                        virtualMachineMenu.MenuItems.Add(saveMenuStateItem);

                        virtualMachineMenu.MenuItems.Add(new MenuItem("-"));
                        if (isPaused)
                        {
                            // We have at least one machine that is paused, so we can set this flag to true.
                            isAnyPaused = true;

                            // Resume
                            var resumeMenuItem = new MenuItem(ResourceHelper.Command_Resume, (_, _) => ControlVirtualMachine(virtualMachineName, VirtualMachineState.Resuming));
                            virtualMachineMenu.MenuItems.Add(resumeMenuItem);
                        }
                        else
                        {
                            // Pause
                            var pauseMenuItem = new MenuItem(ResourceHelper.Command_Pause, (_, _) => ControlVirtualMachine(virtualMachineName, VirtualMachineState.Quiesce));
                            virtualMachineMenu.MenuItems.Add(pauseMenuItem);
                        }

                        // Reset
                        var resetMenuItem = new MenuItem(ResourceHelper.Command_Reset, (_, _) => ControlVirtualMachine(virtualMachineName, VirtualMachineState.Reset));
                        virtualMachineMenu.MenuItems.Add(resetMenuItem);
                    }

                    // Add VM menu item to root menu.
                    ContextMenu.MenuItems.Add(virtualMachineMenu);
                }
            }

            // If there are more than 2 virtual machines, we also create an "All Virtual Machines" menu.
            if (virtualMachines.Count > 1)
            {
                ContextMenu.MenuItems.Add(new MenuItem("-"));

                // Create a root menu item for the `All Virtual Machines` entry.
                var vmItem = new MenuItem(ResourceHelper.Menu_AllVirtualMachines);

                var subItems = new List<MenuItem>();

                // Start
                if (isAnyOffOrSaved)
                {
                    var startAllMenuItem = new MenuItem(ResourceHelper.Command_Start, (_, _) => ControlAllVirtualMachines(VirtualMachineState.Enabled));
                    subItems.Add(startAllMenuItem);
                }

                // Turn Off
                if (isAnyRunning || isAnyPaused)
                {
                    var stopAllMenuItem = new MenuItem(ResourceHelper.Command_TurnOff, (_, _) => ControlAllVirtualMachines(VirtualMachineState.Disabled));
                    subItems.Add(stopAllMenuItem);
                }

                // Shut Down
                if (isAnyRunning)
                {
                    var shutDownAllMenuItem = new MenuItem(ResourceHelper.Command_ShutDown, (_, _) => ControlAllVirtualMachines(VirtualMachineState.ShutDown));
                    subItems.Add(shutDownAllMenuItem);
                }

                // Save
                if (isAnyRunning || isAnyPaused)
                {
                    var saveAllMenuItem = new MenuItem(ResourceHelper.Command_Save, (_, _) => ControlAllVirtualMachines(VirtualMachineState.Offline));
                    subItems.Add(saveAllMenuItem);
                }

                if (subItems.Any() && isAnyRunning || isAnyPaused)
                {
                    subItems.Add(new MenuItem("-"));
                }

                // Pause
                if (isAnyRunning)
                {
                    var pauseAllMenuItem = new MenuItem(ResourceHelper.Command_Pause, (_, _) => ControlAllVirtualMachines(VirtualMachineState.Quiesce));
                    subItems.Add(pauseAllMenuItem);
                }

                // Resume
                if (isAnyPaused)
                {
                    var resumeAllMenuItem = new MenuItem(ResourceHelper.Command_Resume, (_, _) => ControlAllVirtualMachines(VirtualMachineState.Resuming));
                    subItems.Add(resumeAllMenuItem);
                }

                // Reset
                if (isAnyRunning || isAnyPaused)
                {
                    var resetAllMenuItem = new MenuItem(ResourceHelper.Command_Reset, (_, _) => ControlAllVirtualMachines(VirtualMachineState.Reset));
                    subItems.Add(resetAllMenuItem);
                }

                vmItem.MenuItems.AddRange(subItems.ToArray());

                // Add the VM to the context menu.
                ContextMenu.MenuItems.Add(vmItem);
            }

            // Add `Hyper-V Manager` menu item.
            if (_vmManagerPath != null)
            {
                // Only generate if we found the tool installed.

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
            // We want this app to be single instance, so we'll use a mutex to enforce this.
            using (new Mutex(true, ApplicationName, out var firstInstance))
            {
                // Check whether this is the first instance, or if an instance already exists.
                if (firstInstance)
                {
                    // This is the first instance, so we'll run the app as normal.

                    // Application setup.
                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);

                    // Show our tray icon.
                    NotifyIcon.Visible = true;
                    SetTrayIcon();

                    // Check whether the Hyper-V Tools folder exists and, if so, grab or load the required resources.
                    _hyperVInstallFolder = @$"{Environment.GetEnvironmentVariable("ProgramFiles")}\Hyper-V\";
                    if (Directory.Exists(_hyperVInstallFolder))
                    {
                        // Set the path to "vmconnect.exe" (if it exists) so we can launch the VM connection application.
                        _vmConnectPath = $@"{Environment.GetEnvironmentVariable("SYSTEMROOT")}\System32\vmconnect.exe";
                        if (!File.Exists(_vmConnectPath))
                        {
                            _vmConnectPath = null;
                        }

                        // Set the path to "virtmgmt.msc" (if it exists) so we can launch Hyper-V manager.
                        _vmManagerPath = $@"{Environment.GetEnvironmentVariable("SYSTEMROOT")}\System32\virtmgmt.msc";
                        if (!File.Exists(_vmManagerPath))
                        {
                            _vmManagerPath = null;
                        }

                        // Tell ResourceHelper to attempt to load Hyper-V Tools resources.
                        ResourceHelper.LoadExternalResources(_hyperVInstallFolder);
                    }

                    // Initialize and hook our helpers as required.
                    HyperVHelper.VirtualMachineStateChanged += HyperVHelper_VirtualMachineStateChanged;

                    // Run the application.
                    Application.Run();
                }
            }
        }
        private static void OpenHyperVManager()
        {
            // We'll attempt to run Hyper-V Manager.
            try
            {
                // Set our process info.
                var processInfo = new ProcessStartInfo
                {
                    FileName = @$"{Environment.GetEnvironmentVariable("SYSTEMROOT")}\System32\mmc.exe",
                    Arguments = _vmManagerPath!,
                    WorkingDirectory = _hyperVInstallFolder,
                    UseShellExecute = true,
                    Verb = @"runas"
                };

                // Start the application.
                Process.Start(processInfo);
            }
            catch
            {
                // If something happened then there isn't much we can do, so we'll just show a generic error message to let the user know
                // that we couldn't open Hyper-V Manager.
                //ShowError(ResourceHelper.Message_OpenVMConnectFailed);
            }
        }
        private static void SetTrayIcon()
        {
            // Set the tray icon to the correct icon for the size reported by Windows.
            NotifyIcon.Icon = new Icon(ResourceHelper.Icon_HyperV, SystemInformation.SmallIconSize);
            Debug.WriteLine($"Icon size: {SystemInformation.SmallIconSize.Width}x{SystemInformation.SmallIconSize.Height}");
        }
        private static void ShowError(string heading, string text = "")
        {
            // We'll show an error message to the user using a Task Dialog.

            // Set up the Task Dialog using our standard settings for an error.
            var taskDialogPage = new TaskDialogPage
            {
                AllowCancel = false,
                AllowMinimize = false,
                Buttons = new() { TaskDialogButton.Close },
                Caption = AppDomain.CurrentDomain.FriendlyName,
                Heading = heading,
                Icon = TaskDialogIcon.Error,
                Text = text
            };

            // Show the Task Dialog.
            TaskDialog.ShowDialog(taskDialogPage);
        }
        private static void ShowToast(string virtualMachineName, VirtualMachineState state)
        {
            // We'll show an OS toast/tray notification to the user. This varies depending on the OS version, but Windows will handle this
            // for us.

            // Get a friendly display string for the VM state.
            var stateString = HyperVHelper.VmStateToString(state);

            // Determine whether we are reporting a critical state.
            var isCritical = HyperVHelper.IsCriticalState(state);

            // If reporting a critical state, vary the displayed message accordingly.
            if (isCritical)
            {
                stateString = $"{ResourceHelper.Toast_CriticalState}\n{stateString}";
            }

            // Show our toast/notification.
            NotifyIcon.ShowBalloonTip(BalloonTipTimeout, virtualMachineName, stateString, isCritical ? ToolTipIcon.Error : ToolTipIcon.Info);
        }

        #endregion

        #endregion
    }
}