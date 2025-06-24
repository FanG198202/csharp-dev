using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Net;
using System.Web;
using System.IO;
using Ude; // 新增: 用於偵測編碼

namespace EmEditorFloatingMenu
{
    public partial class FloatingMenuPlugin : Form
    {
        private NotifyIcon trayIcon;
        private ContextMenuStrip trayMenu;
        private Timer selectionTimer;
        private Panel floatingPanel;
        private Button btnWebSearch;
        private Button btnTranslate;
        private Button btnCopy;
        private Button btnClose;
        private string selectedText = "";
        private bool isFloatingVisible = false;

        // Windows API 宣告
        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

        [DllImport("user32.dll")]
        static extern bool GetWindowRect(IntPtr hwnd, ref Rectangle rectangle);

        [DllImport("user32.dll")]
        static extern IntPtr GetActiveWindow();

        [DllImport("user32.dll")]
        static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOZORDER = 0x0004;
        private const uint SWP_SHOWWINDOW = 0x0040;
        private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);

        public FloatingMenuPlugin()
        {
            InitializeComponent();
            InitializeTrayIcon();
            InitializeFloatingMenu();
            InitializeTimer();

            // 隱藏主視窗
            this.WindowState = FormWindowState.Minimized;
            this.ShowInTaskbar = false;
            this.Visible = false;
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // 
            // FloatingMenuPlugin
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.Name = "FloatingMenuPlugin";
            this.Text = "EmEditor 浮動功能表";
            this.Load += new System.EventHandler(this.FloatingMenuPlugin_Load);
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FloatingMenuPlugin_FormClosing);

            this.ResumeLayout(false);
        }

        private void InitializeTrayIcon()
        {
            // 建立系統托盤圖示
            trayIcon = new NotifyIcon();
            trayIcon.Icon = SystemIcons.Application;
            trayIcon.Text = "EmEditor 浮動功能表";
            trayIcon.Visible = true;

            // 建立右鍵選單
            trayMenu = new ContextMenuStrip();

            ToolStripMenuItem showItem = new ToolStripMenuItem("顯示");
            showItem.Click += ShowItem_Click;
            trayMenu.Items.Add(showItem);

            trayMenu.Items.Add(new ToolStripSeparator());

            ToolStripMenuItem exitItem = new ToolStripMenuItem("結束程式");
            exitItem.Click += ExitItem_Click;
            trayMenu.Items.Add(exitItem);

            trayIcon.ContextMenuStrip = trayMenu;
            trayIcon.DoubleClick += TrayIcon_DoubleClick;
        }

        private void InitializeFloatingMenu()
        {
            // 建立浮動面板
            floatingPanel = new Panel();
            floatingPanel.Size = new Size(320, 50);
            floatingPanel.BackColor = Color.FromArgb(240, 240, 240);
            floatingPanel.BorderStyle = BorderStyle.FixedSingle;
            floatingPanel.Visible = false;

            // 網路搜索按鈕
            btnWebSearch = new Button();
            btnWebSearch.Text = "網路搜索";
            btnWebSearch.Size = new Size(75, 30);
            btnWebSearch.Location = new Point(5, 10);
            btnWebSearch.Click += BtnWebSearch_Click;
            floatingPanel.Controls.Add(btnWebSearch);

            // 翻譯按鈕
            btnTranslate = new Button();
            btnTranslate.Text = "翻譯";
            btnTranslate.Size = new Size(75, 30);
            btnTranslate.Location = new Point(85, 10);
            btnTranslate.Click += BtnTranslate_Click;
            floatingPanel.Controls.Add(btnTranslate);

            // 複製按鈕
            btnCopy = new Button();
            btnCopy.Text = "複製";
            btnCopy.Size = new Size(75, 30);
            btnCopy.Location = new Point(165, 10);
            btnCopy.Click += BtnCopy_Click;
            floatingPanel.Controls.Add(btnCopy);

            // 關閉按鈕
            btnClose = new Button();
            btnClose.Text = "關閉";
            btnClose.Size = new Size(65, 30);
            btnClose.Location = new Point(245, 10);
            btnClose.Click += BtnClose_Click;
            floatingPanel.Controls.Add(btnClose);

            this.Controls.Add(floatingPanel);
        }

        private void InitializeTimer()
        {
            selectionTimer = new Timer();
            selectionTimer.Interval = 500; // 每500毫秒檢查一次
            selectionTimer.Tick += SelectionTimer_Tick;
            selectionTimer.Start();
        }

        private void SelectionTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                // 檢查當前活動視窗是否為EmEditor
                IntPtr activeWindow = GetForegroundWindow();
                StringBuilder windowTitle = new StringBuilder(256);
                GetWindowText(activeWindow, windowTitle, 256);

                if (windowTitle.ToString().Contains("EmEditor") || windowTitle.ToString().Contains("emeditor"))
                {
                    // 檢查剪貼簿內容變化（簡化的選取檢測）
                    string currentClipboard = "";
                    try
                    {
                        if (Clipboard.ContainsText())
                        {
                            currentClipboard = Clipboard.GetText();
                        }
                    }
                    catch
                    {
                        // 剪貼簿可能被其他程式佔用
                        return;
                    }

                    // 模擬檢測文字選取（實際應該使用EmEditor API）
                    if (!string.IsNullOrEmpty(currentClipboard) && currentClipboard != selectedText && currentClipboard.Length > 0 && currentClipboard.Length < 1000)
                    {
                        selectedText = currentClipboard;
                        ShowFloatingMenu(activeWindow);
                    }
                    else if (string.IsNullOrEmpty(currentClipboard) && isFloatingVisible)
                    {
                        HideFloatingMenu();
                    }
                }
                else if (isFloatingVisible)
                {
                    HideFloatingMenu();
                }
            }
            catch (Exception ex)
            {
                // 記錄錯誤但不中斷程式執行
                System.Diagnostics.Debug.WriteLine("Timer error: " + ex.Message);
            }
        }

        private void ShowFloatingMenu(IntPtr emEditorWindow)
        {
            try
            {
                Rectangle windowRect = new Rectangle();
                GetWindowRect(emEditorWindow, ref windowRect);

                // 計算浮動面板位置（在視窗上方）
                int x = windowRect.Left + (windowRect.Width - floatingPanel.Width) / 2;
                int y = windowRect.Top - floatingPanel.Height - 10;

                // 確保不會超出螢幕範圍
                if (x < 0) x = 0;
                if (y < 0) y = windowRect.Top + 50;

                floatingPanel.Location = new Point(x, y);
                floatingPanel.Visible = true;
                floatingPanel.BringToFront();

                // 設置為最上層
                SetWindowPos(this.Handle, HWND_TOPMOST, x, y, 0, 0, SWP_NOSIZE | SWP_SHOWWINDOW);

                isFloatingVisible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("顯示浮動選單時發生錯誤: " + ex.Message);
            }
        }

        private void HideFloatingMenu()
        {
            floatingPanel.Visible = false;
            isFloatingVisible = false;
        }

        private void BtnWebSearch_Click(object sender, EventArgs e)
        {
            try
            {
                if (!string.IsNullOrEmpty(selectedText))
                {
                    // 使用預設瀏覽器進行Google搜索
                    string searchUrl = "https://www.google.com/search?q=" + HttpUtility.UrlEncode(selectedText);
                    Process.Start(searchUrl);
                }
                HideFloatingMenu();
            }
            catch (Exception ex)
            {
                MessageBox.Show("網路搜索時發生錯誤: " + ex.Message);
            }
        }

        private void BtnTranslate_Click(object sender, EventArgs e)
        {
            try
            {
                if (!string.IsNullOrEmpty(selectedText))
                {
                    // 使用Google翻譯
                    string translateUrl = "https://translate.google.com/?sl=auto&tl=zh-TW&text=" + HttpUtility.UrlEncode(selectedText);
                    Process.Start(translateUrl);
                }
                HideFloatingMenu();
            }
            catch (Exception ex)
            {
                MessageBox.Show("翻譯時發生錯誤: " + ex.Message);
            }
        }

        private void BtnCopy_Click(object sender, EventArgs e)
        {
            try
            {
                if (!string.IsNullOrEmpty(selectedText))
                {
                    Clipboard.SetText(selectedText);
                    MessageBox.Show("已複製到剪貼簿", "複製成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                HideFloatingMenu();
            }
            catch (Exception ex)
            {
                MessageBox.Show("複製時發生錯誤: " + ex.Message);
            }
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            HideFloatingMenu();
            this.Hide();
        }

        private void TrayIcon_DoubleClick(object sender, EventArgs e)
        {
            ShowMainWindow();
        }

        private void ShowItem_Click(object sender, EventArgs e)
        {
            ShowMainWindow();
        }

        private void ExitItem_Click(object sender, EventArgs e)
        {
            ExitApplication();
        }

        private void ShowMainWindow()
        {
            this.Show();
            this.WindowState = FormWindowState.Normal;
            this.ShowInTaskbar = true;
            this.BringToFront();
        }

        private void ExitApplication()
        {
            selectionTimer.Stop();
            trayIcon.Visible = false;
            Application.Exit();
        }

        private void FloatingMenuPlugin_Load(object sender, EventArgs e)
        {
            // 程式啟動時隱藏到系統托盤
            this.Hide();
        }

        private void FloatingMenuPlugin_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 阻止視窗關閉，改為隱藏到系統托盤
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                this.Hide();
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (selectionTimer != null)
                {
                    selectionTimer.Stop();
                    selectionTimer.Dispose();
                }
                if (trayIcon != null)
                {
                    trayIcon.Dispose();
                }
                if (trayMenu != null)
                {
                    trayMenu.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        // ▼▼▼ Ude自動偵測編碼及I/O包裝 ▼▼▼

        /// <summary>
        /// 自動偵測檔案編碼
        /// </summary>
        public static Encoding DetectEncoding(string filePath)
        {
            using (FileStream fs = File.OpenRead(filePath))
            {
                CharsetDetector detector = new CharsetDetector();
                detector.Feed(fs);
                detector.DataEnd();
                if (detector.Charset != null)
                {
                    try
                    {
                        return Encoding.GetEncoding(detector.Charset);
                    }
                    catch
                    {
                        return Encoding.UTF8; // 若無法解析則用UTF8
                    }
                }
                else
                {
                    return Encoding.Default; // 偵測失敗則用系統預設
                }
            }
        }

        /// <summary>
        /// 自動偵測編碼讀取檔案
        /// </summary>
        public static string ReadFileAutoEncoding(string filePath)
        {
            Encoding encoding = DetectEncoding(filePath);
            return File.ReadAllText(filePath, encoding);
        }

        /// <summary>
        /// 以指定編碼寫入檔案
        /// </summary>
        public static void WriteFileAutoEncoding(string filePath, string content, Encoding encoding)
        {
            File.WriteAllText(filePath, content, encoding);
        }

        // ▲▲▲ Ude自動偵測編碼及I/O包裝 ▲▲▲
    }

    // 程式進入點
    public class Program
    {
        [STAThread]
        static void Main()
        {
            // 讓 Console I/O 預設為UTF-8（如需依來源偵測可在此加判斷）
            Console.InputEncoding = Encoding.UTF8;
            Console.OutputEncoding = Encoding.UTF8;

            // ★範例：自動偵測讀取檔案
            /*
            string filePath = "input.txt";
            string content = FloatingMenuPlugin.ReadFileAutoEncoding(filePath);
            Console.WriteLine("偵測自動編碼後讀取內容：");
            Console.WriteLine(content);
            */

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new FloatingMenuPlugin());
        }
    }
}
