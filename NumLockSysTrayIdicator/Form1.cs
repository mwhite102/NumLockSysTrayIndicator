using System;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

/// <summary>
/// Application that displays an notify icon in the system tray to indicate the status of numlock
/// because my new Dell laptop doesn't have a numlock indicator!
/// Keyboard Hook Code based on https://blogs.msdn.microsoft.com/toub/2006/05/03/low-level-keyboard-hook-in-c/
/// </summary>

namespace NumLockSysTrayIdicator
{

    public partial class MainForm : Form
    {
        #region DLL Imports

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        #endregion

        #region Private Properties

        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private static LowLevelKeyboardProc _proc;
        private static IntPtr _hookID = IntPtr.Zero;

        // Context menu for notify icon
        private ContextMenu _NotifyIconContextMenu = new ContextMenu();
        private MenuItem _ExitMenuItem = new MenuItem();

        #endregion

        #region Constructor

        /// <summary>
        /// Constructor
        /// </summary>
        public MainForm()
        {
            InitializeComponent();
            _proc = HookCallback;
            // Set the keyboard hook
            _hookID = SetHook(_proc);

            // Initialize _NotifyIconContextMenu
            _NotifyIconContextMenu.MenuItems.Add(_ExitMenuItem);

            // Initialize _ExitMenuItem;
            _ExitMenuItem.Index = 0;
            _ExitMenuItem.Text = "E&xit";
            _ExitMenuItem.Click += ExitMenuItem_Click;

            NumLock_NotifyIcon.ContextMenu = _NotifyIconContextMenu;
        }
        
        #endregion

        #region Method Overrides

        /// <summary>
        /// Override OnLoad
        /// </summary>
        /// <param name="e"></param>
        protected override void OnLoad(EventArgs e)
        {
            // Hide the main form
            Visible = false;
            // Don't show the form in the taskbar
            ShowInTaskbar = false;
            // Set initial status
            SetNotifyIconStatus(Control.IsKeyLocked(Keys.NumLock));    

            base.OnLoad(e);
        }

        /// <summary>
        /// Override OnFormClosed
        /// </summary>
        /// <param name="e"></param>
        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            // Unhook the keyboard hook
            UnhookWindowsHookEx(_hookID);         
            base.OnFormClosed(e);
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Sets the keyboard hook
        /// </summary>
        /// <param name="proc">The Hook Callback function</param>
        /// <returns>The IntPtr HookId value</returns>
        private IntPtr SetHook(LowLevelKeyboardProc proc)
        {
            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule)
            {
                return SetWindowsHookEx(WH_KEYBOARD_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
            }
        }

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        /// <summary>
        /// The keyboard hook callback
        /// </summary>
        /// <param name="nCode"></param>
        /// <param name="wParam"></param>
        /// <param name="lParam"></param>
        /// <returns></returns>
        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && wParam == (IntPtr)WM_KEYDOWN)
            {
                int vkCode = Marshal.ReadInt32(lParam);
                // Has the NumLock key been pressed
                if ((Keys)vkCode == Keys.NumLock)
                {
                    SetNotifyIconStatus(!Control.IsKeyLocked(Keys.NumLock));
                }
            }
            return CallNextHookEx(_hookID, nCode, wParam, lParam);
        }
        
        /// <summary>
        /// Updates the notify icon text and icon
        /// </summary>
        /// <param name="isNumLockOn">True if numlock is on, false if numlock is off</param>
        private void SetNotifyIconStatus(bool isNumLockOn)
        {
            if (isNumLockOn)
            {
                NumLock_NotifyIcon.Icon = new Icon("NumLockOn.ico");
                NumLock_NotifyIcon.Text = "Numlock Is On";
            }
            else
            {
                NumLock_NotifyIcon.Icon = new Icon("NumLockOff.ico");
                NumLock_NotifyIcon.Text = "Numlock Is Off";
            }
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            // Hide and dispose the notify icon
            NumLock_NotifyIcon.Visible = false;
            NumLock_NotifyIcon.Dispose();
        }

        private void ExitMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        #endregion
    }
}
