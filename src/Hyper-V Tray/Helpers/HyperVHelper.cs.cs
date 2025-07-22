using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;

namespace HyperVTray.Helpers
{
    /// <summary>
    /// Helper class for working with Hyper-V.
    /// </summary>
    internal static class HyperVHelper
    {
        #region Constants

        private const int WmiRefreshInterval = 2;

        #endregion

        #region Fields

        private static readonly ManagementEventWatcher WmiEventWatcher;
        private static readonly ManagementScope WmiManagementScope;

        #endregion

        #region Events

        public static event EventHandler<VirtualMachineStateChangedEventArgs>? VirtualMachineStateChanged;

        #endregion

        #region Construction

        static HyperVHelper()
        {
            // Determine whether we are going to use the v1 or v2 Hyper-V WMI Provider. We use v2 for any version of Windows greater than
            // 6.2 (i.e., Windows 8.1 and Windows Server 2012 R2 onwards).
            var rootWmiPath = Environment.OSVersion.Version > new Version(6, 2) ? "ROOT\\virtualization\\v2" : "ROOT\\virtualization";

            // Create a ManagementScope object, which defines the WMI path we are going to use for our WMI Event Watcher.
            WmiManagementScope = new ManagementScope(rootWmiPath);

            // Set up our WMI Event Watcher to monitor for changes in state to any Hyper-V virtual machine.
            var query = new WqlEventQuery($"SELECT * FROM __InstanceOperationEvent WITHIN {WmiRefreshInterval} " +
                                          $"WHERE TargetInstance ISA 'Msvm_ComputerSystem' " +
                                          $"AND TargetInstance.EnabledState <> PreviousInstance.EnabledState");
            WmiEventWatcher = new ManagementEventWatcher(query) { Scope = WmiManagementScope };
            WmiEventWatcher.EventArrived += WmiEventWatcher_EventArrived;
        }

        #endregion

        #region Methods

        #region Event Handlers

        private static void WmiEventWatcher_EventArrived(object? sender, EventArrivedEventArgs e)
        {
            if (e.NewEvent["TargetInstance"] is ManagementBaseObject virtualMachine)
            {
                var vmState = (VmState)(ushort)virtualMachine["EnabledState"];
                var vmName = virtualMachine["ElementName"].ToString() ?? ResourceHelper.String_UnknownVirtualMachine;

                // Filter out "noisy" states, to just show important ones.
                switch (vmState)
                {
                    case VmState.Enabled:
                    case VmState.Disabled:
                    case VmState.Offline:
                    case VmState.Quiesce:
                    case VmState.Reset:
                        VirtualMachineStateChanged?.Invoke(null, new(vmName, vmState, false));
                        break;

                    case VmState.RunningCritical:
                    case VmState.OffCritical:
                    case VmState.SavedCritical:
                    case VmState.PausedCritical:
                    case VmState.ResetCritical:
                    case VmState.FastSavedCritical:
                        VirtualMachineStateChanged?.Invoke(null, new(vmName, vmState, true));
                        break;
                }
            }
        }

        #endregion

        #region Public Static

        public static IList<ManagementObject> GetVirtualMachines(string? name = null)
        {
            // Create our WMI query string to get one virtual machine, or all virtual machines, from the list of machines configured
            // on this system. If the 'name' parameter is null then we get all virtual machines, otherwise we look for a machine with
            // a matching name.
            var queryString = $"SELECT * FROM Msvm_ComputerSystem WHERE Caption LIKE 'Virtual Machine'";
            if (name != null)
            {
                queryString += $" AND ElementName='{name}'";
            }

            // Create a WMI query object using the querystring we defined previously.
            var queryObj = new ObjectQuery(queryString);

            // Create a searcher object, passing in our search parameters as previously defined.
            var vmSearcher = new ManagementObjectSearcher(WmiManagementScope, queryObj);

            // Run the search, and cast the results to a list of ManagementObject.
            var list = vmSearcher.Get().Cast<ManagementObject>().OrderBy(vm => vm["ElementName"]).ToList();

            // Return the sorted list.
            return list;
        }
        public static void Initialize()
        {
            // Connect to our WMI Management scope.
            WmiManagementScope.Connect();

            // Start our WMI Event Watcher so we receive events on virtual machine state changes.
            WmiEventWatcher.Start();
        }
        public static bool IsVmOff(VmState state)
        {
            // Determine if the VmState enum value is for an off state.
            return state switch
            {
                VmState.Disabled or VmState.OffCritical => true,
                _ => false,
            };
        }
        public static bool IsVmPaused(VmState state)
        {
            // Determine if the VmState enum value is for a paused state.
            return state switch
            {
                VmState.Paused or VmState.Quiesce or VmState.PausedCritical => true,
                _ => false,
            };
        }
        public static bool IsVmRunning(VmState state)
        {
            // Determine if the VmState enum value is for a running state.
            return state switch
            {
                VmState.Enabled or VmState.RunningCritical => true,
                _ => false,
            };
        }
        public static bool IsVmSaved(VmState state)
        {
            // Determine if the VmState enum value is for a saved state.
            return state switch
            {
                VmState.Suspended or VmState.Offline or VmState.SavedCritical or VmState.FastSaved or VmState.FastSavedCritical => true,
                _ => false,
            };
        }
        public static bool RequestVmShutdown(string virtualMachineName)
        {
            // Get the WMI Management Object for the virtual machine we are interested in.
            var virtualMachine = GetVirtualMachines(virtualMachineName).FirstOrDefault();

            // Check to see whether we found a matching virtual machine.
            if (virtualMachine != null)
            {
                // If we found a matching virtual machine get the `Msvm_ShutdownComponent` for it.
                var shutdownComponent = virtualMachine.GetRelated("Msvm_ShutdownComponent").Cast<ManagementObject>().FirstOrDefault();

                // Check to see whether we found a Shutdown Component
                if (shutdownComponent != null)
                {
                    // Get the parameters for the InitiateShutdown method
                    var inParameters = shutdownComponent.GetMethodParameters("InitiateShutdown");

                    // Set the 'Force' and 'Reason' parameters.
                    inParameters["Force"] = true;
                    inParameters["Reason"] = Program.ApplicationName;

                    // Invoke the method, passing in the parameters we just set.
                    var outParameters = shutdownComponent.InvokeMethod("InitiateShutdown", inParameters, null);

                    // Return the result of the method call.
                    if (outParameters != null)
                    {
                        return (StateChangeResponse)outParameters["ReturnValue"] is StateChangeResponse.CompletedwithNoError
                                                                                 or StateChangeResponse.MethodParametersCheckedTransitionStarted;
                    }
                }
            }

            // If we couldn't find a matching virtual machine then return `false`.
            return false;
        }
        public static bool RequestVmStateChange(string virtualMachineName, VmState state)
        {
            // Get the WMI Management Object for the virtual machine we are interested in.
            var virtualMachine = GetVirtualMachines(virtualMachineName).FirstOrDefault();

            // Check to see whether we found a matching virtual machine.
            if (virtualMachine != null)
            {
                // Get the parameters for the 'RequestStateChange' method.
                var inParameters = virtualMachine.GetMethodParameters("RequestStateChange");

                // Filter out the request as we only support a subset requesting a subset of all states.
                if (state is VmState.Enabled  // Running
                          or VmState.Disabled // Stopped
                          or VmState.Offline  // Saved
                          or VmState.Quiesce  // Paused
                          or VmState.Reset)
                {
                    // Set the 'RequestedState' parameter to the desired state.
                    inParameters["RequestedState"] = (ushort)state;

                    // Fire off the request to change the state.
                    var outParameters = virtualMachine.InvokeMethod("RequestStateChange", inParameters, null);

                    // Return the result of the method call.
                    if (outParameters != null)
                    {
                        return (StateChangeResponse)outParameters["ReturnValue"] is StateChangeResponse.CompletedwithNoError
                                                                                 or StateChangeResponse.MethodParametersCheckedTransitionStarted;
                    }
                }
            }

            // If we couldn't find a matching virtual machine then return `false`.
            return false;
        }
        public static string VmStateToString(VmState state)
        {
            // Return the state of a VM as a friendly string.

            switch (state)
            {
                case VmState.Enabled:
                    return ResourceHelper.State_Running;
                case VmState.Offline:
                case VmState.FastSaved:
                    return ResourceHelper.State_Saved;
                case VmState.Quiesce:
                    return ResourceHelper.State_Paused;
                default:
                    if ((int)state >= 32781 && (int)state <= 32792)
                    {
                        return ResourceHelper.State_Critical;
                    }

                    return state.ToString();
            }
        }

        #endregion

        #endregion
    }
}