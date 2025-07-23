using System;

namespace HyperVTray
{
    /// <summary>
    /// EventArgs for when a virtual machine changes state.
    /// </summary>
    internal sealed class VirtualMachineStateChangedEventArgs : EventArgs
    {
        #region Construction

        /// <summary>
        /// Creates a new instance of <see cref="VirtualMachineStateChangedEventArgs"/>.
        /// </summary>
        /// <param name="name">The name of the virtual machine.</param>
        /// <param name="state">The state of the virtual machine.</param>
        public VirtualMachineStateChangedEventArgs(string name, VirtualMachineState state)
        {
            Name = name;
            State = state;
        }

        #endregion

        #region Properties

        /// <summary>
        /// The name of the virtual machine.
        /// </summary>
        public string Name { get; }
        /// <summary>
        /// The state of the virtual machine.
        /// </summary>
        public VirtualMachineState State { get; }

        #endregion
    }
}