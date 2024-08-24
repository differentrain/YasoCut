using System;
using System.Runtime.InteropServices;
using System.Security;
using System.Windows.Input;

namespace YasoCut.PInvoke
{
    [SuppressUnmanagedCodeSecurity]
    internal static class NativeMethods
    {
        [DllImport("Kernel32.dll", CallingConvention = CallingConvention.Winapi, CharSet = CharSet.Unicode, EntryPoint = "GlobalAddAtomW", ExactSpelling = true)]
        public static extern ushort GlobalAddAtom(string lpString);

        [DllImport("Kernel32.dll", CallingConvention = CallingConvention.Winapi, EntryPoint = "GlobalDeleteAtom", ExactSpelling = true)]
        public static extern ushort GlobalDeleteAtom(ushort atom);

        [DllImport("User32.dll", CallingConvention = CallingConvention.Winapi, EntryPoint = "RegisterHotKey", ExactSpelling = true)]
        public static extern bool RegisterHotKey(IntPtr hWnd, int id, ModifierKeys fsModifiers, int vk);

        [DllImport("User32.dll", CallingConvention = CallingConvention.Winapi, EntryPoint = "UnregisterHotKey", ExactSpelling = true)]
        public static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        [DllImport("user32.dll", CallingConvention = CallingConvention.Winapi, CharSet = CharSet.Auto, EntryPoint = "GetForegroundWindow", ExactSpelling = true)]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", EntryPoint = "GetWindowRect", CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern bool GetWindowRect(IntPtr hWnd, ref NativeRect rect);

        [DllImport("user32.dll", CallingConvention = CallingConvention.Winapi, EntryPoint = "GetClientRect", ExactSpelling = true)]
        public static extern bool GetClientRect(IntPtr hWnd, ref NativeRect rect);

        [DllImport("Shell32.dll", CallingConvention = CallingConvention.Winapi, EntryPoint = "SHQueryUserNotificationState", ExactSpelling = true)]
        public static extern int SHQueryUserNotificationState(out int result);

        [DllImport("Dwmapi.dll", CallingConvention = CallingConvention.Winapi, EntryPoint = "DwmGetWindowAttribute", ExactSpelling = true)]
        public static extern int DwmGetWindowAttribute(IntPtr hwnd, int dwAttribute, ref NativeRect rect, int cbAttribute);

        [DllImport("Dwmapi.dll", CallingConvention = CallingConvention.Winapi, EntryPoint = "DwmGetWindowAttribute", ExactSpelling = true)]
        public static extern int DwmGetWindowAttribute(IntPtr hwnd, int dwAttribute, ref int dwMNCRP, int cbAttribute);

        [DllImport("Dwmapi.dll", CallingConvention = CallingConvention.Winapi, EntryPoint = "DwmSetWindowAttribute", ExactSpelling = true)]
        public static extern int DwmSetWindowAttribute(IntPtr hwnd, int dwAttribute, ref int dwMNCRP, int cbAttribute);

        [DllImport("User32.dll", CallingConvention = CallingConvention.Winapi, EntryPoint = "GetWindowDisplayAffinity", ExactSpelling = true)]
        public static extern bool GetWindowDisplayAffinity(IntPtr hwnd, out int pdwAffinity);

        [DllImport("User32.dll", CallingConvention = CallingConvention.Winapi, EntryPoint = "SetWindowDisplayAffinity", ExactSpelling = true)]
        public static extern bool SetWindowDisplayAffinity(IntPtr hwnd, int pdwAffinity);

    }
}
