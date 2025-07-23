namespace HyperVTray
{
    /// <summary>
    /// Represents the possible states a Hyper-V Virtual Machine can be in or requested to go to. This list is taken from both the v1 and
    /// v2 WMI providers, and consists of all states that can be set or returned.
    /// </summary>
    internal enum VirtualMachineState : ushort
    {
        /// <summary>
        /// The state of the element could not be determined.
        /// </summary>
        Unknown = 0,

        Other = 1,

        /// <summary>
        /// The element is running.
        /// </summary>
        Enabled = 2,

        /// <summary>
        /// The element is turned off.
        /// </summary>
        Disabled = 3,

        /// <summary>
        /// The element is in the process of going to a Disabled state.
        /// </summary>
        ShutDown = 4,

        /// <summary>
        /// The element does not support being enabled or disabled.
        /// </summary>
        NotApplicable = 5,

        /// <summary>
        /// The element might be completing commands, and it will drop any new requests.
        /// </summary>
        Offline = 6,

        /// <summary>
        /// The element is in a test state.
        /// </summary>
        Test = 7,

        /// <summary>
        /// The element might be completing commands, but it will queue any new requests.
        /// </summary>
        Defer = 8,

        /// <summary>
        /// The element is enabled but in a restricted mode.The behavior of the element is similar to the Enabled state(2), but it
        /// processes only a restricted set of commands.All other requests are queued.
        /// </summary>
        Quiesce = 9,

        /// <summary>
        /// The element is in the process of going to an Enabled state(2). New requests are queued.
        /// </summary>
        RebootOrStarting = 10,

        /// <summary>
        /// Reset the virtual machine. Corresponds to CIM_EnabledLogicalElement.EnabledState = Reset.
        /// </summary>
        Reset = 11,

        /// <summary>
        /// VM is paused.
        /// </summary>
        Paused = 32768,

        /// <summary>
        /// VM is in a saved state.
        /// </summary>
        Suspended = 32769,

        /// <summary>
        /// VM is starting.
        /// </summary>
        Starting = 32770,

        /// <summary>
        /// In version 1 (V1) of Hyper-V, corresponds to EnabledStateSaving.
        /// </summary>
        Saving = 32773,

        /// <summary>
        /// VM is turning off.
        /// </summary>
        Stopping = 32774,

        /// <summary>
        /// In version 1 (V1) of Hyper-V, corresponds to EnabledStatePausing.
        /// </summary>
        Pausing = 32776,

        /// <summary>
        /// In version 1 (V1) of Hyper-V, corresponds to EnabledStateResuming. State transition from Paused to Running.
        /// </summary>
        Resuming = 32777,

        /// <summary>
        /// Corresponds to EnabledStateFastSuspend.
        /// </summary>
        FastSaved = 32779,

        /// <summary>
        ///  Corresponds to EnabledStateFastSuspending. State transition from Running to FastSaved.
        /// </summary>
        FastSaving = 32780,

        // The following values represent critical states:
        RunningCritical = 32781,
        OffCritical = 32782,
        StoppingCritical = 32783,
        SavedCritical = 32784,
        PausedCritical = 32785,
        StartingCritical = 32786,
        ResetCritical = 32787,
        SavingCritical = 32788,
        PausingCritical = 32789,
        ResumingCritical = 32790,
        FastSavedCritical = 32791,
        FastSavingCritical = 32792,
    }
}