using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Text;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Timer = System.Windows.Forms.Timer;

namespace WpmMeter;

public partial class MainForm : Form {
    private const int ROLLING_WINDOW_SECONDS = 12;
    private const int KEYSTROKES_PER_WORD = 5;

    private double lastWpm = 0;
    private long totalKeystrokes = 0;

    private static readonly string KeystrokesFilePath = Path.Combine(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "WpmMeter"), "keystrokes.bin");

    private readonly Queue<DateTime> _keystrokeTimestamps = new();
    private readonly ContextMenuStrip trayMenu = new();
    private readonly NotifyIcon notifyIcon = new() { Visible = true, Text = "WPM: 0\nKeystrokes: 0" };

    // Set the font and color.
    private readonly Font font = new("Tahoma", 10, FontStyle.Bold, GraphicsUnit.Point);
    private readonly SolidBrush brush = new(Color.FromArgb(70, 106, 70));

    // Timer to periodically save keystrokes.
    private readonly Timer saveTimer;

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

        // Load persisted keystrokes (if present) before updating the tray icon.
        LoadKeystrokes();
        UpdateTrayIcon(0);

        // Start periodic save timer (every minute).
        saveTimer = new System.Windows.Forms.Timer { Interval = 60_000 };
        saveTimer.Tick += (s, e) => SaveKeystrokes();
        saveTimer.Start();

        // Subscribe to global key events.
        WindowsKeyboardHook.OnGlobalKey += OnGlobalKey;
        FormClosed += (s, e) => {
            // Ensure keystrokes are saved when the form closes.
            SaveKeystrokes();

            WindowsKeyboardHook.OnGlobalKey -= OnGlobalKey;
            notifyIcon.Visible = false;
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
        Interlocked.Increment(ref totalKeystrokes);

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
        if (Math.Abs(Math.Round(wpm) - Math.Round(lastWpm)) > 0) {
            UpdateTrayIcon((int)Math.Round(wpm));
            lastWpm = wpm;
        }
    }

    private void UpdateTrayIcon(int wpm) {
        // Update the tooltip text.
        notifyIcon.Text = $"WPM: {wpm}\nKeystrokes: {totalKeystrokes:N0}";

        // Get the width based on number of digits in the WPM.
        var digitCount = Math.Max(2, wpm.ToString().Length);
        var imageWidth = 16 + (digitCount - 2) * 8;

        // Create a bitmap and draw the WPM text.
        using var bmp = new Bitmap(imageWidth, 16);
        using (var gfx = Graphics.FromImage(bmp)) {
            gfx.Clear(Color.Transparent);
            gfx.TextRenderingHint = TextRenderingHint.SingleBitPerPixelGridFit;

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
        if (imageWidth > 16) {
            // Scale the entire 24x16 image down to 16x16.
            using var gfx = Graphics.FromImage(squeezedBmp);
            gfx.Clear(Color.Transparent);
            gfx.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
            gfx.DrawImage(bmp, new Rectangle(0, 0, 16, 16), new Rectangle(0, 0, imageWidth, 16), GraphicsUnit.Pixel);
        }

        // Create the final icon bitmap, scaling if necessary.
        using var iconBmp = imageWidth > 16
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
            var newIcon = (Icon)icon.Clone();
            var oldIcon = notifyIcon.Icon;
            notifyIcon.Icon = newIcon;
            oldIcon?.Dispose();
        } finally {
            NativeMethods.DestroyIcon(hIcon);
        }
    }

    private static partial class NativeMethods {
        [LibraryImport("user32.dll", EntryPoint = "DestroyIcon")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static partial bool DestroyIcon(IntPtr handle);
    }

    protected override void Dispose(bool disposing) {
        if (disposing) {
            // Ensure OnGlobalKey is unsubscribed.
            WindowsKeyboardHook.OnGlobalKey -= OnGlobalKey;

            // Ensure keystrokes are persisted.
            SaveKeystrokes();

            // Stop and dispose the timer.
            saveTimer?.Stop();
            saveTimer?.Dispose();

            // Dispose the icon currently assigned to the NotifyIcon first to free its GDI handles.
            notifyIcon.Visible = false;
            var currentIcon = notifyIcon.Icon;
            notifyIcon.Icon = null;
            currentIcon?.Dispose();

            // Dispose remaining assets.
            notifyIcon.Dispose();
            trayMenu.Dispose();
            font.Dispose();
            brush.Dispose();
        }

        base.Dispose(disposing);
    }

    private void SaveKeystrokes() {
        var path = KeystrokesFilePath;
        var dir = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(dir)) {
            Directory.CreateDirectory(dir);
        }

        // Read the current value atomically and write it as a binary Int64.
        var value = Interlocked.Read(ref totalKeystrokes);
        using var fs = File.Open(path, FileMode.Create, FileAccess.Write, FileShare.None);
        using var bw = new BinaryWriter(fs);
        bw.Write(value);
    }

    private void LoadKeystrokes() {
        var path = KeystrokesFilePath;
        if (!File.Exists(path)) {
            return;
        }

        using var fs = File.OpenRead(path);
        using var br = new BinaryReader(fs);
        var value = br.ReadInt64();
        Interlocked.Exchange(ref totalKeystrokes, value);

        // Update the tray tooltip to reflect the loaded keystrokes.
        UpdateTrayIcon((int)Math.Round(lastWpm));
    }
}
