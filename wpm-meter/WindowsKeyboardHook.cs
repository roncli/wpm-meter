using System;
using System.Runtime.InteropServices;

namespace WpmMeter;

public static partial class WindowsKeyboardHook {
    private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

    private const int WH_KEYBOARD_LL = 13;
    private const int WM_KEYDOWN = 0x0100;

    private static IntPtr hookPtr = IntPtr.Zero;
    private static LowLevelKeyboardProc hookCall;

    public static event Action OnGlobalKey;

    [LibraryImport("user32.dll", EntryPoint = "SetWindowsHookExW", SetLastError = true)]
    private static partial IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

    [LibraryImport("user32.dll", EntryPoint = "UnhookWindowsHookEx", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool UnhookWindowsHookEx(IntPtr hhk);

    [LibraryImport("user32.dll", EntryPoint = "CallNextHookEx", SetLastError = true)]
    private static partial IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

    public static void EnableHook() {
        hookCall = new LowLevelKeyboardProc(HookCallback);

        hookPtr = SetWindowsHookEx(WH_KEYBOARD_LL, hookCall, IntPtr.Zero, 0);
    }

    public static void DisableHook() {
        UnhookWindowsHookEx(hookPtr);
    }

    private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam) {
        if (nCode >= 0 && wParam == WM_KEYDOWN) {
            var vkCode = Marshal.ReadInt32(lParam);

            if (
                vkCode is not 8 // Back
                and not 16 // ShiftKey
                and not 160 // LShiftKey
                and not 161 // RShiftKey
                and not 65536 // Shift
                and not 262144 // Alt
                and not 131072 // Control
                and not 17 // ControlKey
                and not 162 // LControlKey
                and not 163 // RControlKey
                and not 20 // CapsLock
                and not 91 // LWin
                and not 92 // RWin
            ) {
                OnGlobalKey.Invoke();
            }
        }

        return CallNextHookEx(hookPtr, nCode, wParam, lParam);
    }
}
