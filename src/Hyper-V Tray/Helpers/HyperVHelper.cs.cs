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

            // Connect to our WMI Management scope.
            WmiManagementScope.Connect();

            // Set up our WMI Event Watcher to monitor for changes in state to any Hyper-V virtual machine.
            var query = new WqlEventQuery($"SELECT * FROM __InstanceOperationEvent WITHIN {WmiRefreshInterval} " +
                                          $"WHERE TargetInstance ISA 'Msvm_ComputerSystem' " +
                                          $"AND TargetInstance.EnabledState <> PreviousInstance.EnabledState");
            WmiEventWatcher = new ManagementEventWatcher(query) { Scope = WmiManagementScope };
            WmiEventWatcher.EventArrived += WmiEventWatcher_EventArrived;

            // Start our WMI Event Watcher so we receive events on virtual machine state changes.
            WmiEventWatcher.Start();
        }

        #endregion

        #region Methods

        #region Event Handlers

        private static void WmiEventWatcher_EventArrived(object? sender, EventArrivedEventArgs e)
        {
            // We've had a WMI event from our watcher - let's check it's what we want.
            if (e.NewEvent["TargetInstance"] is ManagementBaseObject virtualMachine)
            {
                // We've had WMI event for a VM, so we'll check it's new state and raise our own event so that the tray icon can respond.

                // Grab the VM state and name.
                var vmState = (VirtualMachineState)(ushort)virtualMachine["EnabledState"];
                var vmName = virtualMachine["ElementName"].ToString() ?? ResourceHelper.String_UnknownVirtualMachine;

                // Filter out "noisy" states, to just raise events for "important" ones.
                if (vmState is VirtualMachineState.Disabled
                            or VirtualMachineState.Enabled
                            or VirtualMachineState.FastSavedCritical
                            or VirtualMachineState.OffCritical
                            or VirtualMachineState.Offline
                            or VirtualMachineState.PausedCritical
                            or VirtualMachineState.Quiesce
                            or VirtualMachineState.Reset
                            or VirtualMachineState.ResetCritical
                            or VirtualMachineState.RunningCritical
                            or VirtualMachineState.SavedCritical)
                    VirtualMachineStateChanged?.Invoke(null, new(vmName, vmState));
            }
        }

        #endregion

        #region Public Static

        /// <summary>
        /// Gets a list of one or more virtual machines via WMI. If <see cref="name"/> name is <see langword="null"/> then all machines are returned; otherwise, just
        /// the first matching virtual machine is returned.
        /// </summary>
        /// <param name="name">The name of a virtual machine to get.</param>
        /// <returns>A <see cref="List{T}"/> of 0 or more <see cref="ManagementObject"/> instances, representing the matching virtual
        /// machines.</returns>
        public static IList<ManagementObject> GetVirtualMachines(string? name = null)
        {
            // Create our WMI query string to get one virtual machine, or all virtual machines, from the list of machines configured
            // on this system. If the 'name' parameter is null then we get all virtual machines, otherwise we look for a machine with
            // a matching name.
            var queryString = "SELECT * FROM Msvm_ComputerSystem WHERE Caption LIKE 'Virtual Machine'";
            if (name != null)
            {
                queryString += $" AND ElementName='{name}'";
            }

            // Create a WMI query object using the querystring we defined previously.
            var queryObj = new ObjectQuery(queryString);

            // Create a searcher object, passing in our search parameters as previously defined.
            var vmSearcher = new ManagementObjectSearcher(WmiManagementScope, queryObj);

            // Run the search, and cast the results to an ordered list of ManagementObject.
            var list = vmSearcher.Get().Cast<ManagementObject>().OrderBy(vm => vm["ElementName"]).ToList();

            // Return the list.
            return list;
        }
        /// <summary>
        /// Returns a <see cref="bool"/> indicating whether a <see cref="VirtualMachineState"/> is a critical state.
        /// </summary>
        /// <param name="state">The <see cref="VirtualMachineState"/> to check.</param>
        /// <returns><see langword="true"/> if the state is a critical state; otherwise <see langword="false"/>.</returns>
        public static bool IsCriticalState(VirtualMachineState state)
        {
            // Determine if the VmState enum value is a critical state.
            var stateValue = (int)state;
            return stateValue is >= 32781 and <= 32792;
        }
        /// <summary>
        /// Returns a <see cref="bool"/> indicating whether a <see cref="VirtualMachineState"/> is a powered off state.
        /// </summary>
        /// <param name="state">The <see cref="VirtualMachineState"/> to check.</param>
        /// <returns><see langword="true"/> if the state is a powered off state; otherwise <see langword="false"/>.</returns>
        public static bool IsOffState(VirtualMachineState state)
        {
            // Determine if the VmState enum value is for an off state.
            return state switch
            {
                VirtualMachineState.Disabled or 
                VirtualMachineState.OffCritical => true,
                _ => false,
            };
        }
        /// <summary>
        /// Returns a <see cref="bool"/> indicating whether a <see cref="VirtualMachineState"/> is a paused state.
        /// </summary>
        /// <param name="state">The <see cref="VirtualMachineState"/> to check.</param>
        /// <returns><see langword="true"/> if the state is a paused state; otherwise <see langword="false"/>.</returns>
        public static bool IsPausedState(VirtualMachineState state)
        {
            // Determine if the VmState enum value is for a paused state.
            return state switch
            {
                VirtualMachineState.Paused or
                VirtualMachineState.Quiesce or
                VirtualMachineState.PausedCritical => true,
                _ => false,
            };
        }
        /// <summary>
        /// Returns a <see cref="bool"/> indicating whether a <see cref="VirtualMachineState"/> is a running state.
        /// </summary>
        /// <param name="state">The <see cref="VirtualMachineState"/> to check.</param>
        /// <returns><see langword="true"/> if the state is a running state; otherwise <see langword="false"/>.</returns>
        public static bool IsRunningState(VirtualMachineState state)
        {
            // Determine if the VmState enum value is for a running state.
            return state switch
            {
                VirtualMachineState.Enabled or
                VirtualMachineState.RunningCritical => true,
                _ => false,
            };
        }
        /// <summary>
        /// Returns a <see cref="bool"/> indicating whether a <see cref="VirtualMachineState"/> is a saved state.
        /// </summary>
        /// <param name="state">The <see cref="VirtualMachineState"/> to check.</param>
        /// <returns><see langword="true"/> if the state is a saved state; otherwise <see langword="false"/>.</returns>
        public static bool IsSavedState(VirtualMachineState state)
        {
            // Determine if the VmState enum value is for a saved state.
            return state switch
            {
                VirtualMachineState.Suspended or
                VirtualMachineState.Offline or
                VirtualMachineState.SavedCritical or
                VirtualMachineState.FastSaved or
                VirtualMachineState.FastSavedCritical => true,
                _ => false,
            };
        }
        /// <summary>
        /// Requests the specified virtual machine to perform a graceful shutdown.
        /// </summary>
        /// <param name="virtualMachineName">The name of the virtual machine to shutdown.</param>
        /// <returns><see langword="true"/> if the virtual machine was successfully requested to shutdown; otherwise
        /// <see langword="false"/>.</returns>
        public static bool RequestVirtualMachineShutdown(string virtualMachineName)
        {
            // Flag to indicate whether the shutdown requests was a success.
            var result = false;

            // Get the WMI Management Object for the virtual machine we are interested in.
            var virtualMachine = GetVirtualMachines(virtualMachineName).FirstOrDefault();

            // If we found a matching virtual machine get the `Msvm_ShutdownComponent` for it.
            var shutdownComponent = virtualMachine?.GetRelated("Msvm_ShutdownComponent").Cast<ManagementObject>().FirstOrDefault();

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
                    result = (StateChangeResponse)outParameters["ReturnValue"] is StateChangeResponse.CompletedwithNoError
                        or StateChangeResponse.MethodParametersCheckedTransitionStarted;
                }
            }

            // Return our result.
            return result;
        }
        /// <summary>
        /// Requests the specified virtual machine to change state.
        /// </summary>
        /// <param name="virtualMachineName">The name of the virtual machine to change state.</param>
        /// <param name="state">The <see cref="VirtualMachineState"/> the virtual machine should transition to.</param>
        /// <returns><see langword="true"/> if the virtual machine was successfully requested to change state; otherwise
        /// <see langword="false"/>.</returns>
        public static bool RequestVirtualMachineStateChange(string virtualMachineName, VirtualMachineState state)
        {
            // Flag to indicate whether the state change was a success.
            var result = false;

            // Get the WMI Management Object for the virtual machine we are interested in.
            var virtualMachine = GetVirtualMachines(virtualMachineName).FirstOrDefault();

            // Check to see whether we found a matching virtual machine.
            if (virtualMachine != null)
            {
                // Get the parameters for the 'RequestStateChange' method.
                var inParameters = virtualMachine.GetMethodParameters("RequestStateChange");

                // Get the VM state.
                var virtualMachineState = (VirtualMachineState)Convert.ToInt32(virtualMachine["EnabledState"]);
                
                // Filter out the request as we only support a subset requesting a subset of all states.
                if (virtualMachineState != state && state is VirtualMachineState.Enabled  // Running
                                                          or VirtualMachineState.Disabled // Stopped
                                                          or VirtualMachineState.Offline  // Saved
                                                          or VirtualMachineState.Quiesce  // Paused
                                                          or VirtualMachineState.Reset)
                {
                    // Set the 'RequestedState' parameter to the desired state.
                    inParameters["RequestedState"] = (ushort)state;

                    // Fire off the request to change the state.
                    var outParameters = virtualMachine.InvokeMethod("RequestStateChange", inParameters, null);

                    // Set our result flag to the result of the method call.
                    if (outParameters != null)
                    {
                        result = (StateChangeResponse)outParameters["ReturnValue"] is StateChangeResponse.CompletedwithNoError
                                                                                   or StateChangeResponse.MethodParametersCheckedTransitionStarted
                                                                                   or StateChangeResponse.InvalidStateForThisOperation;
                    }
                }
                else
                {
                    // The VM is already in the requested state, so we will set our success flag to true.
                    result = true;
                }
            }

            // Return our result.
            return result;
        }
        /// <summary>
        /// Returns a friendly display string for a specified <see cref="VirtualMachineState"/>.
        /// </summary>
        /// <param name="state">The <see cref="VirtualMachineState"/> to get a display string for.</param>
        /// <returns>A <see cref="string"/> containing the friendly display value.</returns>
        public static string VmStateToString(VirtualMachineState state)
        {
            // Get the state of a VM as a friendly string.
            var stateString = state switch
            {
                VirtualMachineState.Disabled => ResourceHelper.State_Off,
                VirtualMachineState.Enabled => ResourceHelper.State_Running,
                VirtualMachineState.Offline => ResourceHelper.State_Saved,
                VirtualMachineState.FastSaved => ResourceHelper.State_Saved,
                VirtualMachineState.Quiesce => ResourceHelper.State_Paused,
                _ => ResourceHelper.State_Unknown
            };

            // If we're critical, format the state string appropriately.
            if (IsCriticalState(state))
            {
                stateString = string.Format(ResourceHelper.State_Critical, stateString);
            }

            // Return the state string.
            return stateString;
        }

        #endregion

        #endregion
    }
}