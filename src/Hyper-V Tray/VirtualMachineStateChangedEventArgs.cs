using System;

namespace HyperVTray
{
    internal sealed class VirtualMachineStateChangedEventArgs : EventArgs
    {
        #region Construction

        public VirtualMachineStateChangedEventArgs(string name, VmState state, bool critical)
        {
            Critical = critical;
            Name = name;
            State = state;
        }

        #endregion

        #region Properties

        public bool Critical { get; }
        public string Name { get; }
        public VmState State { get; }

        #endregion
    }
}