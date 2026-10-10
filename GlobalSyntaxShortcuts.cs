using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace VsNetSyntaxWordAddIn
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public sealed class ShortcutDiagnostics
    {
        private readonly Func<string> status;
        internal ShortcutDiagnostics(Func<string> status) { this.status = status; }
        public string Status => status();
    }

    internal sealed class GlobalSyntaxShortcuts : NativeWindow, IDisposable
    {
        private const uint Modifiers = 0x0001 | 0x0002 | 0x0004 | 0x4000;
        private static readonly uint[] Keys = { 0x43, 0x4A, 0x54, 0x48 };
        private readonly bool[] registered = new bool[Keys.Length];
        private readonly Action<string> apply;
        private readonly Func<bool> wordActive;
        private readonly uint processId = (uint)Process.GetCurrentProcess().Id;
        private readonly Timer focusTimer;
        private readonly ForegroundCallback foregroundCallback;
        private readonly IntPtr foregroundHook;
        private bool disposed;
        private bool applying;
        private int lastError;
        private int dispatched;

        internal int RegisteredCount
        {
            get
            {
                int count = 0;
                foreach (bool value in registered) if (value) count++;
                return count;
            }
        }

        internal string Status => "Registered=" + RegisteredCount + "/4; Dispatched=" + dispatched +
            "; LastError=" + lastError;

        internal GlobalSyntaxShortcuts(Action<string> apply, Func<bool> wordActive = null)
        {
            this.apply = apply ?? throw new ArgumentNullException(nameof(apply));
            this.wordActive = wordActive ?? IsWordActive;
            CreateHandle(new CreateParams { Caption = "VS.NET Syntax Shortcuts", Parent = new IntPtr(-3) });
            foregroundCallback = OnForegroundChanged;
            foregroundHook = SetWinEventHook(3, 3, IntPtr.Zero, foregroundCallback, 0, 0, 0);
            // Retry registration after focus changes or another Word instance releases a shortcut.
            focusTimer = new Timer { Interval = 100 };
            focusTimer.Tick += (sender, args) => Refresh();
            Refresh();
            focusTimer.Start();
        }

        internal void Refresh()
        {
            if (disposed) return;
            bool active = wordActive();
            lastError = 0;
            for (int i = 0; i < Keys.Length; i++)
            {
                if (active && !registered[i])
                {
                    registered[i] = RegisterHotKey(Handle, (int)Keys[i], Modifiers, Keys[i]);
                    if (!registered[i]) lastError = Marshal.GetLastWin32Error();
                }
                else if (!active && registered[i])
                {
                    UnregisterHotKey(Handle, (int)Keys[i]);
                    registered[i] = false;
                }
            }
        }

        protected override void WndProc(ref Message message)
        {
            if (message.Msg == 0x0312 && !disposed)
            {
                int index = Array.IndexOf(Keys, (uint)message.WParam.ToInt32());
                if (index >= 0 && registered[index] && wordActive() && !applying)
                {
                    applying = true;
                    try
                    {
                        dispatched++;
                        apply(LanguageForKey(Keys[index]));
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "VS.NET Syntax", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    finally { applying = false; }
                }
                return;
            }
            base.WndProc(ref message);
        }

        private static string LanguageForKey(uint key)
        {
            switch (key)
            {
                case 0x43: return "CSharp";
                case 0x4A: return "JavaScript";
                case 0x54: return "TypeScript";
                case 0x48: return "Html";
                default: throw new ArgumentOutOfRangeException(nameof(key));
            }
        }

        private bool IsWordActive()
        {
            IntPtr window = GetForegroundWindow();
            uint owner;
            GetWindowThreadProcessId(window, out owner);
            if (owner != processId) return false;
            var name = new StringBuilder(32);
            GetClassName(window, name, name.Capacity);
            return name.ToString() == "OpusApp";
        }

        private void OnForegroundChanged(IntPtr hook, uint eventType, IntPtr window,
            int objectId, int childId, uint threadId, uint time) => Refresh();

        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            focusTimer.Stop();
            focusTimer.Dispose();
            if (foregroundHook != IntPtr.Zero) UnhookWinEvent(foregroundHook);
            for (int i = 0; i < Keys.Length; i++)
            {
                if (registered[i]) UnregisterHotKey(Handle, (int)Keys[i]);
                registered[i] = false;
            }
            DestroyHandle();
        }

        private delegate void ForegroundCallback(IntPtr hook, uint eventType, IntPtr window,
            int objectId, int childId, uint threadId, uint time);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool RegisterHotKey(IntPtr window, int id, uint modifiers, uint key);
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnregisterHotKey(IntPtr window, int id);
        [DllImport("user32.dll")]
        private static extern IntPtr SetWinEventHook(uint min, uint max, IntPtr module,
            ForegroundCallback callback, uint processId, uint threadId, uint flags);
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWinEvent(IntPtr hook);
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr window, StringBuilder name, int length);
    }
}
