using System;
using System.ComponentModel;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using System.Windows.Forms;

namespace DotNetControls
{
    /// <summary>
    /// This interface is used as the COM Source interface for the ExampleButton control class.
    /// </summary>
    /// <remarks>
    /// <para>The COM source interface allows for use of the COM connection points protocol.</para>
    /// <para>All events that this control needs to expose should be defined in this
    /// interface as VB6 only supports its WithEvents syntax for a single interface.</para>
    /// <para>Each method defined in this interface must match up to an event in the user
    /// control having the same name; the method signatures here must match the signature of
    /// the corresponding event's delegate.</para>
    /// <para>Interface is defined as a dispinterface (IDispatch) because VB6 requires it
    /// for source interfaces.</para>
    /// <para>Each method must be marked with a unique DispId value greater than 0. Without proper
    /// DispIds, raising an event may cause a COMException to be thrown if the VB6 client does not
    /// handle all defined events.</para>
    /// </remarks>
    [ComVisible(true), Guid(ExampleButton.EventsId), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ExampleButtonInteropEvents
    {
        [DispId(1)]
        void Click();
    }

    /// <summary>
    /// This is the default interface implemented by the user control, and should
    /// contain all the methods and properties that will be exposed to COM.
    /// </summary>
    [ComVisible(true), Guid(ExampleButton.InterfaceId)]
    public interface IExampleButton
    {
        /// <summary>
        /// Gets or sets a value indicating whether the user control is visible.
        /// </summary>
        [DispId(1)]
        bool Visible { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the user control is enabled.
        /// </summary>
        [DispId(2)]
        bool Enabled { get; set; }

        /// <summary>
        /// Gets or sets the foreground color of the user control.
        /// </summary>
        [DispId(3)]
        int ForegroundColor { get; set; }

        /// <summary>
        /// Gets or sets the background color of the user control.
        /// </summary>
        [DispId(4)]
        int BackgroundColor { get; set; }

        /// <summary>
        /// Gets or sets the background image of the user control.
        /// </summary>
        [DispId(5)]
        Image BackgroundImage { get; set; }

        /// <summary>
        /// Forces the control to invalidate its client area and immediately redraw 
        /// itself and any child controls.
        /// </summary>
        [DispId(6)]
        void Refresh();
    }

    [ComVisible(true),
        Guid(ClassId),
        ClassInterface(ClassInterfaceType.None),
        ComSourceInterfaces(typeof(ExampleButtonInteropEvents))]
    public partial class ExampleButton: UserControl, IExampleButton
    {
        #region VB6 Interop Code
        #region VB6 Properties

        // The following are examples of how to expose typical form properties to VB6.  
        // You can also use these as examples on how to add additional properties.

        // Must declare this property as new as it exists in Windows.Forms and is not overridable
        public new bool Visible
        {
            get { return base.Visible; }
            set { base.Visible = value; }
        }

        public new bool Enabled
        {
            get { return base.Enabled; }
            set { base.Enabled = value; }
        }

        public int ForegroundColor
        {
            get
            {
                return ActiveXControlHelpers.GetOleColorFromColor(base.ForeColor);
            }
            set
            {
                base.ForeColor = ActiveXControlHelpers.GetColorFromOleColor(value);
            }
        }

        public int BackgroundColor
        {
            get
            {
                return ActiveXControlHelpers.GetOleColorFromColor(base.BackColor);
            }
            set
            {
                base.BackColor = ActiveXControlHelpers.GetColorFromOleColor(value);
            }
        }

        public override Image BackgroundImage
        {
            get { return null; }
            set
            {
                if (null != value)
                {
                    MessageBox.Show("Setting the background image of an Interop UserControl is not supported, please use a PictureBox instead.", "Information");
                }
                base.BackgroundImage = null;
            }
        }

        #endregion

        // These routines perform the additional COM registration needed by ActiveX controls
        [EditorBrowsable(EditorBrowsableState.Never)]
        [ComRegisterFunction]
        private static void Register(System.Type t)
        {
            ComRegistration.RegisterControl(t);
        }

        [EditorBrowsable(EditorBrowsableState.Never)]
        [ComUnregisterFunction]
        private static void Unregister(System.Type t)
        {
            ComRegistration.UnregisterControl(t);
        }

        [SecurityPermission(SecurityAction.LinkDemand, Flags = SecurityPermissionFlag.UnmanagedCode)]
        protected override void WndProc(ref Message m)
        {

            const int WM_SETFOCUS = 0x7;
            const int WM_PARENTNOTIFY = 0x210;
            const int WM_DESTROY = 0x2;
            const int WM_LBUTTONDOWN = 0x201;
            const int WM_RBUTTONDOWN = 0x204;

            if (m.Msg == WM_SETFOCUS)
            {
                // Raise Enter event
                this.OnEnter(System.EventArgs.Empty);
            }
            else if (m.Msg == WM_PARENTNOTIFY && (m.WParam.ToInt32() == WM_LBUTTONDOWN || m.WParam.ToInt32() == WM_RBUTTONDOWN))
            {

                if (!this.ContainsFocus)
                {
                    // Raise Enter event
                    this.OnEnter(System.EventArgs.Empty);
                }
            }
            else if (m.Msg == WM_DESTROY && !this.IsDisposed && !this.Disposing)
            {
                // Used to ensure that VB6 will cleanup control properly
                this.Dispose();
            }

            base.WndProc(ref m);
        }

        // Ensures that the Validating and Validated events fire appropriately
        internal void ValidationHandler(object sender, System.EventArgs e)
        {
            if (this.ContainsFocus) return;

            // Raise Leave event
            this.OnLeave(e);

            if (this.CausesValidation)
            {
                CancelEventArgs validationArgs = new CancelEventArgs();
                this.OnValidating(validationArgs);

                if (validationArgs.Cancel && this.ActiveControl != null)
                    this.ActiveControl.Focus();
                else
                {
                    // Raise Validated event
                    this.OnValidated(e);
                }
            }

        }

        // This event will hook up the necessary handlers
        protected virtual void ControlAddedHandlerForInterop(object sender, ControlEventArgs e)
        {
            ActiveXControlHelpers.WireUpHandlers(e.Control, ValidationHandler);
        }
        // Ensures that tabbing across VB6 and .NET controls works as expected
        protected virtual void LostFocusHandlerForInterop(object sender, System.EventArgs e)
        {
            ActiveXControlHelpers.HandleFocus(this);
        }
        #endregion


        public const string ClassId = "93C9D89E-5D96-4A1A-A07B-144478408B4B";
        public const string InterfaceId = "A10D1930-BC8A-4574-BFAB-20F612027128";
        public const string EventsId = "84A58A0A-6BAF-4FC4-816E-6CA52B698482";

        
        public delegate void ClickHandler();
        public new event ClickHandler Click; // use "new" because Click is already defined by the base class

        public ExampleButton()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Click?.Invoke();
        }
    }
}
