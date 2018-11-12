using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Runtime.InteropServices;


namespace XL.Office.Helpers
{
    public delegate int KeyHandlerDelegate(bool repeated);

    public struct KeyState
    {
        public KeyState(Keys key, bool ctrl = false, bool alt = false, bool shift = false)
        {
            Key = key;
            Ctrl = ctrl;
            Alt = alt;
            Shift = shift;
        }

        public Keys Key;
        public bool Ctrl;
        public bool Alt;
        public bool Shift;
    }

    public class InterceptKeys
    {
        public delegate int LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
        private static LowLevelKeyboardProc _proc = HookCallback;
        private static IntPtr _hookID = IntPtr.Zero;

        private const int WH_KEYBOARD = 2;
        private const int HC_ACTION = 0;

        private static Dictionary<KeyState, KeyHandlerDelegate> KeyHandlers;

        public static void SetHooks(Dictionary<KeyState, KeyHandlerDelegate> handlers)
        {
            if (handlers == null) return;
            KeyHandlers = handlers;
#pragma warning disable 618
            _hookID = SetWindowsHookEx(WH_KEYBOARD, _proc, IntPtr.Zero, (uint)AppDomain.GetCurrentThreadId());
#pragma warning restore 618
        }

        public static void ReleaseHook()
        {
            UnhookWindowsHookEx(_hookID);
        }

        private static int HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            int PreviousStateBit = 31;
            bool KeyRepeated = false;

            Int64 bitmask = (Int64)Math.Pow(2, (PreviousStateBit - 1));

            try
            {

                if (nCode < 0)
                {
                    return (int)CallNextHookEx(_hookID, nCode, wParam, lParam);
                }
                else
                {
                    if (nCode == HC_ACTION)
                    {
                        KeyRepeated = ((Int64)lParam & bitmask) > 0;
                        Keys keyData = (Keys)wParam;
                        var keys = new KeyState(keyData, ctrl: IsKeyDown(Keys.ControlKey), alt: IsKeyDown(Keys.Menu), shift: IsKeyDown(Keys.ShiftKey));
                        if (KeyHandlers.ContainsKey(keys))
                        {
                            return KeyHandlers[keys](KeyRepeated);
                        }
                    }
                    return (int)CallNextHookEx(_hookID, nCode, wParam, lParam);
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message, "InterceptKeys Exception:");
                return (int)CallNextHookEx(_hookID, nCode, wParam, lParam);
            }
        }

        public static bool IsKeyDown(Keys keys)
        {
            return (GetKeyState((int)keys) & 0x8000) == 0x8000;
        }

        [DllImport("user32.dll")]
        private static extern short GetKeyState(int nVirtKey);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode,
            IntPtr wParam, IntPtr lParam);
    }
}
