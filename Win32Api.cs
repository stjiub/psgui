using System;
using System.Runtime.InteropServices;

public class Win32 {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

    public static IntPtr FindWindowByProcessId(int processId) {
        IntPtr foundHandle = IntPtr.Zero;

        EnumWindows(delegate (IntPtr hWnd, IntPtr lParam) {
            int windowProcessId;
            GetWindowThreadProcessId(hWnd, out windowProcessId);
            if (windowProcessId == processId) {
                foundHandle = hWnd;
                return false;  // Stop enumerating
            }
            return true;  // Continue enumerating
        }, IntPtr.Zero);

        return foundHandle;
    }

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    
    [DllImport("user32.dll", SetLastError=true)]
    public static extern bool SetParent(IntPtr hWndChild, IntPtr hWndNewParent);
    
    [DllImport("user32.dll", SetLastError=true)]
    public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

    [DllImport("user32.dll", SetLastError=true)]
    public static extern bool SetFocus(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError=true)]
    public static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll", SetLastError = true)]
    public static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
}