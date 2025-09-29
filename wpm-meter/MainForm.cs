using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Text;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace WpmMeter;

public partial class MainForm : Form {
    private const int ROLLING_WINDOW_SECONDS = 12;
    private const int KEYSTROKES_PER_WORD = 5;

    private double lastWpm = 0;
    private int totalKeystrokes = 0;

    private readonly Queue<DateTime> _keystrokeTimestamps = new();
    private readonly ContextMenuStrip trayMenu = new();
    private readonly NotifyIcon notifyIcon = new() { Visible = true, Text = "WPM: 0\nKeystrokes: 0" };

    public MainForm() {
        // Hide the form.
        ShowInTaskbar = false;
        WindowState = FormWindowState.Minimized;
        Visible = false;
        Opacity = 0;
        Load += (s, e) => Hide();

        // Create context menu for tray icon.
        var exitItem = new ToolStripMenuItem("E&xit") {
            ShowShortcutKeys = true
        };
        exitItem.Click += (s, e) => Application.Exit();
        trayMenu.Items.Add(exitItem);

        // Create the tray icon.
        notifyIcon.ContextMenuStrip = trayMenu;
        UpdateTrayIcon(0);

        // Subscribe to global key events.
        WindowsKeyboardHook.OnGlobalKey += OnGlobalKey;
        FormClosed += (s, e) => {
            WindowsKeyboardHook.OnGlobalKey -= OnGlobalKey;
            notifyIcon.Visible = false;
            notifyIcon.Dispose();
        };
    }

    private void OnGlobalKey() {
        // Ensure we're on the UI thread.
        if (InvokeRequired) {
            BeginInvoke(new Action(OnGlobalKey));
            return;
        }

        // Record the keystroke timestamp and increment the total keystrokes.
        var now = DateTime.UtcNow;
        _keystrokeTimestamps.Enqueue(now);
        totalKeystrokes++;

        // Remove keystrokes older than the rolling window.
        while (_keystrokeTimestamps.Count > 0 && (now - _keystrokeTimestamps.Peek()).TotalSeconds > ROLLING_WINDOW_SECONDS) {
            _keystrokeTimestamps.Dequeue();
        }

        // Calculate WPM based on time between first and last keystroke.
        var wpm = 0d;
        if (_keystrokeTimestamps.Count > 1) {
            var first = _keystrokeTimestamps.Peek();
            var minutes = (now - first).TotalMinutes;
            if (minutes > 0) {
                var words = _keystrokeTimestamps.Count / (double)KEYSTROKES_PER_WORD;
                wpm = words / minutes;
            }
        }

        // Update the tray icon if WPM has changed.
        if (Math.Abs(wpm - lastWpm) > 0.5) {
            UpdateTrayIcon((int)Math.Round(wpm));
            lastWpm = wpm;
        }
    }

    private void UpdateTrayIcon(int wpm) {
        // Update the tooltip text.
        notifyIcon.Text = $"WPM: {wpm}\nKeystrokes: {totalKeystrokes:N0}";

        // Get the width based on WPM.
        var imageWidth = wpm >= 100 ? 24 : 16;

        // Create a bitmap and draw the WPM text.
        using var bmp = new Bitmap(imageWidth, 16);
        using (var gfx = Graphics.FromImage(bmp)) {
            gfx.Clear(Color.Transparent);
            gfx.TextRenderingHint = TextRenderingHint.SingleBitPerPixelGridFit;

            // Set the font and color.
            using var font = new Font("Tahoma", 10, FontStyle.Bold, GraphicsUnit.Point);
            using var brush = new SolidBrush(Color.FromArgb(70, 106, 70));

            // Set the text and measure its size.
            var text = wpm.ToString();
            var size = gfx.MeasureString(text, font);

            // Draw the text centered in the bitmap.
            var x = (bmp.Width - size.Width) / 2f + 1;
            var y = (bmp.Height - size.Height) / 2f + 1;
            gfx.DrawString(text, font, brush, x, y);
        }

        // Create the final icon bitmap, scaling if necessary.
        using var squeezedBmp = new Bitmap(16, 16);
        if (imageWidth == 24) {
            // Scale the entire 24x16 image down to 16x16.
            using var gfx = Graphics.FromImage(squeezedBmp);
            gfx.Clear(Color.Transparent);
            gfx.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
            gfx.DrawImage(bmp, new Rectangle(0, 0, 16, 16), new Rectangle(0, 0, 24, 16), GraphicsUnit.Pixel);
        }

        // Create the final icon bitmap, scaling if necessary.
        using var iconBmp = imageWidth == 24
            ? squeezedBmp.Clone(new Rectangle(0, 0, 16, 16), squeezedBmp.PixelFormat)
            : (Bitmap)bmp.Clone();

        // Convert the bitmap to an icon and set it on notifyIcon.
        IntPtr hIcon;
        try {
            hIcon = iconBmp.GetHicon();
        } catch (Exception) {
            return;
        }

        try {
            using var icon = Icon.FromHandle(hIcon);
            notifyIcon.Icon = (Icon)icon.Clone();
        } finally {
            NativeMethods.DestroyIcon(hIcon);
        }
    }

    private static partial class NativeMethods {
        [LibraryImport("user32.dll", EntryPoint = "DestroyIcon")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static partial bool DestroyIcon(IntPtr handle);
    }
}
