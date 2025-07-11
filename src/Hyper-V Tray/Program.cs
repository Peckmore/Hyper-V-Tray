using HyperVTray.Helpers;
using Microsoft.Toolkit.Uwp.Notifications;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
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

        private static void ExitMenuItem_Click(object? sender, EventArgs e)
        {
            NotifyIcon.Visible = false;
            Application.Exit();
        }
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
        private static void PauseMenuItem_Click(object? sender, EventArgs e)
        {
            if (sender is MenuItem menuItem)
            {
                if (!HyperVHelper.RequestVmStateChange(menuItem.Parent.Name, VmState.Quiesce))
                {
                    ShowError(string.Format(ResourceHelper.Message_PauseVMFailed, menuItem.Parent.Name));
                }
            }
        }
        private static void ResetMenuItem_Click(object? sender, EventArgs e)
        {
            if (sender is MenuItem menuItem)
            {
                if (!HyperVHelper.RequestVmStateChange(menuItem.Parent.Name, VmState.Reset))
                {
                    ShowError(ResourceHelper.Message_ResetVMFailed, menuItem.Parent.Name);
                }
            }
        }
        private static void ResumeMenuItem_Click(object? sender, EventArgs e)
        {
            if (sender is MenuItem menuItem)
            {
                if (!HyperVHelper.RequestVmStateChange(menuItem.Parent.Name, VmState.Enabled))
                {
                    ShowError(string.Format(ResourceHelper.Message_ResumeVMFailed, menuItem.Parent.Name));
                }
            }
        }
        private static void SaveMenuItem_Click(object? sender, EventArgs e)
        {
            if (sender is MenuItem menuItem)
            {
                if (!HyperVHelper.RequestVmStateChange(menuItem.Parent.Name, VmState.Offline))
                {
                    ShowError(string.Format(ResourceHelper.Message_SaveStateVMFailed, menuItem.Parent.Name));
                }
            }
        }
        private static void ShutDownMenuItem_Click(object? sender, EventArgs e)
        {
            if (sender is MenuItem menuItem)
            {
                if (!HyperVHelper.RequestVmShutdown(menuItem.Parent.Name))
                {
                    ShowError(string.Format(ResourceHelper.Message_ShutDownVMFailed, menuItem.Parent.Name));
                }
            }
        }
        private static void StartMenuItem_Click(object? sender, EventArgs e)
        {
            if (sender is MenuItem menuItem)
            {
                if (!HyperVHelper.RequestVmStateChange(menuItem.Parent.Name, VmState.Enabled))
                {
                    ShowError(string.Format(ResourceHelper.Message_StartVMFailed, menuItem.Parent.Name));
                }
            }
        }
        private static void SystemEvents_DisplaySettingsChanged(object? sender, EventArgs e)
        {
            Debug.WriteLine("Display settings change detected.");

            SetTrayIcon();
        }
        private static void TurnOffMenuItem_Click(object? sender, EventArgs e)
        {
            if (sender is MenuItem menuItem)
            {
                if (!HyperVHelper.RequestVmStateChange(menuItem.Parent.Name, VmState.Disabled))
                {
                    ShowError(string.Format(ResourceHelper.Message_PowerOffVMFailed, menuItem.Parent.Name));
                }
            }
        }

        #endregion
            
        #region Private Static

        private static void ConnectToVm(string virtualMachineName)
        {
            var processInfo = new ProcessStartInfo
            {
                FileName = _vmConnectPath!,
                Arguments = $"localhost \"{virtualMachineName}\"",
            };
            Process.Start(processInfo);
        }
        private static void GenerateContextMenu()
        {
            // Clear the context menu.
            ContextMenu.MenuItems.Clear();

            // Get all VMs.
            var virtualMachines = HyperVHelper.GetVirtualMachines().ToList();

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
                    var canResume = HyperVHelper.IsVmPaused(virtualMachineStatus); // Paused
                    var canStart = HyperVHelper.IsVmOff(virtualMachineStatus) || HyperVHelper.IsVmSaved(virtualMachineStatus); // Stopped or Saved

                    // Now generate control menu items for VM.

                    // Connect
                    if (_vmConnectPath != null)
                    {
                        var connectMenuItem = new MenuItem(ResourceHelper.Command_Connect);
                        connectMenuItem.Click += (_, _) => ConnectToVm(virtualMachineName);
                        virtualMachineMenu.MenuItems.Add(connectMenuItem);
                        virtualMachineMenu.MenuItems.Add(new MenuItem("-"));
                    }

                    if (canStart)
                    {
                        // Start
                        var startMenuItem = new MenuItem(ResourceHelper.Command_Start);
                        startMenuItem.Click += StartMenuItem_Click;
                        virtualMachineMenu.MenuItems.Add(startMenuItem);
                    }
                    else
                    {
                        // Turn Off
                        var stopMenuItem = new MenuItem(ResourceHelper.Command_TurnOff);
                        stopMenuItem.Click += TurnOffMenuItem_Click;
                        virtualMachineMenu.MenuItems.Add(stopMenuItem);

                        // Shut Down
                        if (!canResume)
                        {
                            var shutMenuDownItem = new MenuItem(ResourceHelper.Command_ShutDown);
                            shutMenuDownItem.Click += ShutDownMenuItem_Click;
                            virtualMachineMenu.MenuItems.Add(shutMenuDownItem);
                        }

                        // Save
                        var saveMenuStateItem = new MenuItem(ResourceHelper.Command_Save);
                        saveMenuStateItem.Click += SaveMenuItem_Click;
                        virtualMachineMenu.MenuItems.Add(saveMenuStateItem);

                        virtualMachineMenu.MenuItems.Add(new MenuItem("-"));
                        if (canResume)
                        {
                            // Resume
                            var resumeMenuItem = new MenuItem(ResourceHelper.Command_Resume);
                            resumeMenuItem.Click += ResumeMenuItem_Click;
                            virtualMachineMenu.MenuItems.Add(resumeMenuItem);
                        }
                        else
                        {
                            // Pause
                            var pauseMenuItem = new MenuItem(ResourceHelper.Command_Pause);
                            pauseMenuItem.Click += PauseMenuItem_Click;
                            virtualMachineMenu.MenuItems.Add(pauseMenuItem);
                        }

                        // Reset
                        var resetMenuItem = new MenuItem(ResourceHelper.Command_Reset);
                        resetMenuItem.Click += ResetMenuItem_Click;
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
                var isOff = virtualMachines.Any(vm => HyperVHelper.IsVmOff((VmState)Convert.ToInt32(vm["EnabledState"])));
                var isPaused = virtualMachines.Any(vm => HyperVHelper.IsVmPaused((VmState)Convert.ToInt32(vm["EnabledState"])));
                var isRunning = virtualMachines.Any(vm => HyperVHelper.IsVmRunning((VmState)Convert.ToInt32(vm["EnabledState"])));
                var isSaved = virtualMachines.Any(vm => HyperVHelper.IsVmSaved((VmState)Convert.ToInt32(vm["EnabledState"])));

                // Start
                if (isOff || isSaved)
                {
                    subItems.Add(new MenuItem(ResourceHelper.Command_Start)); //MenuItemStartAll_Click));
                }

                // Turn Off
                if (isRunning || isPaused)
                {
                    subItems.Add(new MenuItem(ResourceHelper.Command_TurnOff)); //MenuItemStopAll_Click));
                }

                // Shut Down
                if (isRunning)
                {
                    subItems.Add(new MenuItem(ResourceHelper.Command_ShutDown)); //MenuItemShutdownAll_Click));
                }

                // Save
                if (isRunning || isPaused)
                {
                    subItems.Add(new MenuItem(ResourceHelper.Command_Save)); //MenuItemSaveAll_Click));
                }

                if (subItems.Any() && isRunning || isPaused)
                {
                    subItems.Add(new MenuItem("-"));
                }

                // Resume
                if (isPaused)
                {
                    subItems.Add(new MenuItem(ResourceHelper.Command_Resume)); //MenuItemResumeAll_Click));
                }

                // Pause
                if (isRunning)
                {
                    subItems.Add(new MenuItem(ResourceHelper.Command_Pause)); //MenuItemPauseAll_Click));
                }

                // Reset
                if (isRunning || isPaused)
                {
                    subItems.Add(new MenuItem(ResourceHelper.Command_Reset)); //MenuItemResetAll_Click));
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
            exitItem.Click += ExitMenuItem_Click;
            ContextMenu.MenuItems.Add(exitItem);
        }
        [STAThread]
        private static void Main()
        {
            // Application setup.
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // Detect Hyper-V components.
            _hyperVInstallFolder = @$"{Environment.GetEnvironmentVariable("ProgramFiles")}\Hyper-V\";
            if (!Directory.Exists(_hyperVInstallFolder))
            {
                ShowError("Hyper-V Tools not installed.");
                Application.Exit();
                return;
            }
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

            // Initialize our helpers.
            HyperVHelper.Initialize();
            HyperVHelper.VirtualMachineStateChanged += HyperVHelper_VirtualMachineStateChanged;
            ResourceHelper.Initialize(_hyperVInstallFolder);

            // Show our tray icon.
            NotifyIcon.Visible = true;
            SetTrayIcon();

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
            var window = typeof(NotifyIcon).GetField("_window", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(NotifyIcon);
            var windowHandle = (IntPtr)window!.GetType().GetProperty("Handle")!.GetValue(window)!;
            var iconWidth = PInvoke.GetTrayIconWidth(windowHandle);
            Debug.Assert(iconWidth > 0, "Icon width is 0");
            var iconSize = new Size(iconWidth, iconWidth);
            NotifyIcon.Icon = new Icon(ResourceHelper.Icon_HyperV, iconSize);

            Debug.WriteLine($"Icon size: {iconSize.Width}x{iconSize.Height}");
        }
        private static void ShowError(string heading, string text = "")
        {
            var taskDialogPage = new TaskDialogPage();
            taskDialogPage.AllowCancel = false;
            taskDialogPage.AllowMinimize = false;
            taskDialogPage.Buttons = new() { TaskDialogButton.Close };
            taskDialogPage.Caption = ApplicationName;
            taskDialogPage.Heading = heading;
            taskDialogPage.Icon = TaskDialogIcon.Error;
            taskDialogPage.Text = text;
            TaskDialog.ShowDialog(taskDialogPage);
        }
        private static void ShowToast(string virtualMachineName, VmState vmState, bool isCritical)
        {
            var status = HyperVHelper.VmStateToString(vmState);
            if (Environment.OSVersion.Version.Major >= 10)
            {
                var toast = new ToastContentBuilder().AddHeader(virtualMachineName, virtualMachineName, new ToastArguments());

                if (isCritical)
                { 
                    toast.AddText(ResourceHelper.Toast_CriticalState, AdaptiveTextStyle.Header);
                }

                toast.AddText(status)
                     .Show();
            }
            else
            {
                if (isCritical)
                {
                    status = $"{ResourceHelper.Toast_CriticalState}\n{status}";
                }
                NotifyIcon.ShowBalloonTip(BalloonTipTimeout, virtualMachineName, status, isCritical ? ToolTipIcon.Error : ToolTipIcon.Info);
            }
        }

        #endregion

        #endregion
    }
}