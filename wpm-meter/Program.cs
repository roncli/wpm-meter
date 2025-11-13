using System;
using System.Windows.Forms;

namespace WpmMeter;

internal static class Program {
    [STAThread]
    private static void Main() {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        WindowsKeyboardHook.EnableHook();
        Application.Run(new MainForm());
        WindowsKeyboardHook.DisableHook();
    }
}
