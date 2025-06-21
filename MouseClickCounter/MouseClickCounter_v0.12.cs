using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Diagnostics;
using System.Text;

namespace MouseClickCounter
{
    public partial class Program
    {
        private const int WH_MOUSE_LL = 14;
        private const int WM_LBUTTONDOWN = 0x0201;

        private static LowLevelMouseProc _proc = HookCallback;
        private static IntPtr _hookID = IntPtr.Zero;
        private static NotifyIcon notifyIcon;
        private static string currentMouseModel = "";
        private static Dictionary<string, MouseData> mouseDataDict = new Dictionary<string, MouseData>();
        private static string dataDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "MouseClickCounter");

        public delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelMouseProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        [STAThread]
        public static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            bool createdNew;
            using (Mutex mutex = new Mutex(true, "MouseClickCounterMutex", out createdNew))
            {
                if (!createdNew)
                {
                    MessageBox.Show("程序已在運行中！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                SetAutoStart();

                Directory.CreateDirectory(dataDirectory);

                InitializeNotifyIcon();

                DetectMouseModel();

                _hookID = SetHook(_proc);

                Application.Run();

                UnhookWindowsHookEx(_hookID);
            }
        }

        private static void SetAutoStart()
        {
            try
            {
                RegistryKey key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", true);
                string appPath = Application.ExecutablePath;
                key.SetValue("MouseClickCounter", appPath);
                key.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("設置自動啟動失敗: " + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static void InitializeNotifyIcon()
        {
            notifyIcon = new NotifyIcon();
            notifyIcon.Icon = SystemIcons.Application;
            notifyIcon.Text = "滑鼠點擊次數記錄器";
            notifyIcon.Visible = true;

            ContextMenuStrip contextMenu = new ContextMenuStrip();
            contextMenu.Items.Add("設定", null, OnSettingsClick);
            contextMenu.Items.Add("關於", null, OnAboutClick);
            contextMenu.Items.Add("-");
            contextMenu.Items.Add("離開", null, OnExitClick);

            notifyIcon.ContextMenuStrip = contextMenu;
        }

        private static void DetectMouseModel()
        {
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_PointingDevice");
                foreach (ManagementObject obj in searcher.Get())
                {
                    string manufacturer = obj["Manufacturer"] != null ? obj["Manufacturer"].ToString() : "Unknown";
                    string name = obj["Name"] != null ? obj["Name"].ToString() : "Unknown Mouse";
                    currentMouseModel = manufacturer + " - " + name;
                    break;
                }

                if (string.IsNullOrEmpty(currentMouseModel))
                {
                    currentMouseModel = "Unknown Mouse";
                }

                CheckNewMouseModel();
            }
            catch (Exception ex)
            {
                MessageBox.Show("檢測滑鼠型號失敗: " + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                currentMouseModel = "Unknown Mouse";
            }
        }

        private static void CheckNewMouseModel()
        {
            string filePath = Path.Combine(dataDirectory, SanitizeFileName(currentMouseModel) + ".html");

            if (!mouseDataDict.ContainsKey(currentMouseModel))
            {
                if (File.Exists(filePath))
                {
                    LoadMouseData(currentMouseModel);
                }
                else
                {
                    CreateNewMouseData(currentMouseModel);
                }
            }
        }

        private static void CreateNewMouseData(string mouseModel)
        {
            MouseData data = new MouseData
            {
                MouseModel = mouseModel,
                SwitchBrand = "未設定",
                SwitchModel = "未設定",
                ImagePath = "",
                DailyClicks = 0,
                TotalClicks = 0,
                CreationDate = DateTime.Now,
                LastClickDate = DateTime.Now.Date
            };

            mouseDataDict[mouseModel] = data;
            SaveMouseData(mouseModel);
        }

        private static void LoadMouseData(string mouseModel)
        {
            try
            {
                string filePath = Path.Combine(dataDirectory, SanitizeFileName(mouseModel) + ".html");
                if (File.Exists(filePath))
                {
                    string content = File.ReadAllText(filePath, Encoding.UTF8);
                    MouseData data = ParseHtmlData(content);
                    data.MouseModel = mouseModel;

                    if (data.LastClickDate.Date != DateTime.Now.Date)
                    {
                        data.DailyClicks = 0;
                        data.LastClickDate = DateTime.Now.Date;
                    }

                    mouseDataDict[mouseModel] = data;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("載入滑鼠數據失敗: " + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                CreateNewMouseData(mouseModel);
            }
        }

        private static MouseData ParseHtmlData(string htmlContent)
        {
            MouseData data = new MouseData();

            try
            {
                if (htmlContent.Contains("微動開關品牌："))
                {
                    int start = htmlContent.IndexOf("微動開關品牌：") + "微動開關品牌：".Length;
                    int end = htmlContent.IndexOf("<", start);
                    data.SwitchBrand = htmlContent.Substring(start, end - start).Trim();
                }

                if (htmlContent.Contains("微動開關型號："))
                {
                    int start = htmlContent.IndexOf("微動開關型號：") + "微動開關型號：".Length;
                    int end = htmlContent.IndexOf("<", start);
                    data.SwitchModel = htmlContent.Substring(start, end - start).Trim();
                }

                if (htmlContent.Contains("今日點擊次數："))
                {
                    int start = htmlContent.IndexOf("今日點擊次數：") + "今日點擊次數：".Length;
                    int end = htmlContent.IndexOf("<", start);
                    int tempDailyClicks;
                    if (int.TryParse(htmlContent.Substring(start, end - start).Trim(), out tempDailyClicks))
                    {
                        data.DailyClicks = tempDailyClicks;
                    }
                }

                if (htmlContent.Contains("總點擊次數："))
                {
                    int start = htmlContent.IndexOf("總點擊次數：") + "總點擊次數：".Length;
                    int end = htmlContent.IndexOf("<", start);
                    int tempTotalClicks;
                    if (int.TryParse(htmlContent.Substring(start, end - start).Trim(), out tempTotalClicks))
                    {
                        data.TotalClicks = tempTotalClicks;
                    }
                }

                if (htmlContent.Contains("建立日期："))
                {
                    int start = htmlContent.IndexOf("建立日期：") + "建立日期：".Length;
                    int end = htmlContent.IndexOf("<", start);
                    DateTime tempCreationDate;
                    if (DateTime.TryParse(htmlContent.Substring(start, end - start).Trim(), out tempCreationDate))
                    {
                        data.CreationDate = tempCreationDate;
                    }
                }

                data.LastClickDate = DateTime.Now.Date;
            }
            catch
            {
                // 解析失敗時使用默認值
            }

            return data;
        }

        private static void SaveMouseData(string mouseModel)
        {
            try
            {
                if (!mouseDataDict.ContainsKey(mouseModel)) return;

                MouseData data = mouseDataDict[mouseModel];
                string filePath = Path.Combine(dataDirectory, SanitizeFileName(mouseModel) + ".html");

                string htmlContent = GenerateHtmlReport(data);
                File.WriteAllText(filePath, htmlContent, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                MessageBox.Show("保存滑鼠數據失敗: " + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static string GenerateHtmlReport(MouseData data)
        {
            string imageHtml = "";
            if (!string.IsNullOrEmpty(data.ImagePath) && File.Exists(data.ImagePath))
            {
                imageHtml = "<img src=\"" + data.ImagePath + "\" alt=\"微動開關圖片\" style=\"max-width: 200px; max-height: 150px;\" />";
            }

            return "<!DOCTYPE html>\n" +
                   "<html>\n" +
                   "<head>\n" +
                   "    <meta charset='UTF-8'>\n" +
                   "    <title>滑鼠點擊記錄 - " + data.MouseModel + "</title>\n" +
                   "    <style>\n" +
                   "        body { font-family: Arial, sans-serif; margin: 20px; }\n" +
                   "        .info-table { border-collapse: collapse; width: 100%; }\n" +
                   "        .info-table th, .info-table td { border: 1px solid #ddd; padding: 8px; text-align: left; }\n" +
                   "        .info-table th { background-color: #f2f2f2; }\n" +
                   "        .image-container { text-align: center; margin: 20px 0; }\n" +
                   "    </style>\n" +
                   "</head>\n" +
                   "<body>\n" +
                   "    <h1>滑鼠點擊記錄報告</h1>\n" +
                   "    \n" +
                   "    <table class='info-table'>\n" +
                   "        <tr><th>項目</th><th>資訊</th></tr>\n" +
                   "        <tr><td>滑鼠型號</td><td>" + data.MouseModel + "</td></tr>\n" +
                   "        <tr><td>微動開關品牌</td><td>" + data.SwitchBrand + "</td></tr>\n" +
                   "        <tr><td>微動開關型號</td><td>" + data.SwitchModel + "</td></tr>\n" +
                   "        <tr><td>今日點擊次數</td><td>" + data.DailyClicks + "</td></tr>\n" +
                   "        <tr><td>總點擊次數</td><td>" + data.TotalClicks + "</td></tr>\n" +
                   "        <tr><td>建立日期</td><td>" + data.CreationDate.ToString("yyyy-MM-dd HH:mm:ss") + "</td></tr>\n" +
                   "        <tr><td>最後點擊日期</td><td>" + data.LastClickDate.ToString("yyyy-MM-dd") + "</td></tr>\n" +
                   "    </table>\n" +
                   "    \n" +
                   "    <div class='image-container'>\n" +
                   "        " + imageHtml + "\n" +
                   "    </div>\n" +
                   "    \n" +
                   "    <p><small>最後更新時間: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "</small></p>\n" +
                   "</body>\n" +
                   "</html>";
        }

        private static string SanitizeFileName(string fileName)
        {
            char[] invalidChars = Path.GetInvalidFileNameChars();
            return new string(fileName.Where(c => !invalidChars.Contains(c)).ToArray());
        }

        private static IntPtr SetHook(LowLevelMouseProc proc)
        {
            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule)
            {
                return SetWindowsHookEx(WH_MOUSE_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
            }
        }

        private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && wParam == (IntPtr)WM_LBUTTONDOWN)
            {
                RecordClick();
            }

            return CallNextHookEx(_hookID, nCode, wParam, lParam);
        }

        private static void RecordClick()
        {
            if (!string.IsNullOrEmpty(currentMouseModel) && mouseDataDict.ContainsKey(currentMouseModel))
            {
                MouseData data = mouseDataDict[currentMouseModel];

                if (data.LastClickDate.Date != DateTime.Now.Date)
                {
                    data.DailyClicks = 0;
                    data.LastClickDate = DateTime.Now.Date;
                }

                data.DailyClicks++;
                data.TotalClicks++;

                if (data.TotalClicks % 100 == 0)
                {
                    SaveMouseData(currentMouseModel);
                }
            }
        }

        private static void OnSettingsClick(object sender, EventArgs e)
        {
            SettingsForm settingsForm = new SettingsForm(mouseDataDict, currentMouseModel);
            settingsForm.DataUpdated += delegate(object s, EventArgs args) { SaveMouseData(currentMouseModel); };
            settingsForm.ShowDialog();
        }

        private static void OnAboutClick(object sender, EventArgs e)
        {
            MessageBox.Show(
                "程序名稱：滑鼠點擊次數記錄器\n" +
                "程序版本：v1.0.0\n" +
                "程序功能：記錄當前滑鼠使用的微動開關已點擊次數和理論壽命值對比\n" +
                "程序版權：c 2025 滑鼠點擊記錄器\n" +
                "開發日期：2025年6月",
                "關於", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private static void OnExitClick(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show(
                "確定要結束程序嗎？\n\n" +
                "提醒：不同型號滑鼠及微動開關點擊記錄會分開記錄，\n" +
                "但是分段記錄對精確評估會有誤差值。",
                "確認離開", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                foreach (string mouseModel in mouseDataDict.Keys)
                {
                    SaveMouseData(mouseModel);
                }

                notifyIcon.Visible = false;
                Application.Exit();
            }
        }
    }

    public class MouseData
    {
        public string MouseModel { get; set; }
        public string SwitchBrand { get; set; }
        public string SwitchModel { get; set; }
        public string ImagePath { get; set; }
        public int DailyClicks { get; set; }
        public int TotalClicks { get; set; }
        public DateTime CreationDate { get; set; }
        public DateTime LastClickDate { get; set; }
    }

    public partial class SettingsForm : Form
    {
        private Dictionary<string, MouseData> mouseDataDict;
        private string currentMouseModel;
        private DataGridView dataGridView;
        private Button btnSave, btnReset, btnBrowseImage;
        private TextBox txtImagePath;

        public event EventHandler DataUpdated;

        public SettingsForm(Dictionary<string, MouseData> data, string currentModel)
        {
            mouseDataDict = data;
            currentMouseModel = currentModel;
            InitializeComponent();
            LoadData();
        }

        private void InitializeComponent()
        {
            this.Text = "滑鼠設定";
            this.Size = new Size(800, 600);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;

            dataGridView = new DataGridView();
            dataGridView.Location = new Point(10, 10);
            dataGridView.Size = new Size(760, 400);
            dataGridView.AllowUserToAddRows = false;
            dataGridView.AllowUserToDeleteRows = false;
            dataGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView.MultiSelect = false;

            dataGridView.Columns.Add("MouseModel", "滑鼠型號");
            dataGridView.Columns.Add("SwitchBrand", "微動開關品牌");
            dataGridView.Columns.Add("SwitchModel", "微動開關型號");
            dataGridView.Columns.Add("DailyClicks", "今日點擊");
            dataGridView.Columns.Add("TotalClicks", "總點擊");
            dataGridView.Columns.Add("CreationDate", "創建日期");

            dataGridView.Columns["MouseModel"].ReadOnly = true;
            dataGridView.Columns["DailyClicks"].ReadOnly = true;
            dataGridView.Columns["TotalClicks"].ReadOnly = true;
            dataGridView.Columns["CreationDate"].ReadOnly = true;

            Label lblImage = new Label();
            lblImage.Text = "微動開關圖片：";
            lblImage.Location = new Point(10, 430);
            lblImage.Size = new Size(100, 23);

            txtImagePath = new TextBox();
            txtImagePath.Location = new Point(120, 430);
            txtImagePath.Size = new Size(500, 23);

            btnBrowseImage = new Button();
            btnBrowseImage.Text = "瀏覽";
            btnBrowseImage.Location = new Point(630, 430);
            btnBrowseImage.Size = new Size(75, 23);
            btnBrowseImage.Click += BtnBrowseImage_Click;

            btnSave = new Button();
            btnSave.Text = "保存";
            btnSave.Location = new Point(530, 470);
            btnSave.Size = new Size(75, 30);
            btnSave.Click += BtnSave_Click;

            btnReset = new Button();
            btnReset.Text = "重置記錄";
            btnReset.Location = new Point(620, 470);
            btnReset.Size = new Size(75, 30);
            btnReset.Click += BtnReset_Click;

            Button btnClose = new Button();
            btnClose.Text = "關閉";
            btnClose.Location = new Point(710, 470);
            btnClose.Size = new Size(75, 30);
            btnClose.Click += (s, e) => this.Close();

            this.Controls.AddRange(new Control[] {
                dataGridView, lblImage, txtImagePath, btnBrowseImage,
                btnSave, btnReset, btnClose
            });

            dataGridView.SelectionChanged += DataGridView_SelectionChanged;
        }

        private void LoadData()
        {
            dataGridView.Rows.Clear();
            foreach (var kvp in mouseDataDict)
            {
                MouseData data = kvp.Value;
                dataGridView.Rows.Add(
                    data.MouseModel,
                    data.SwitchBrand,
                    data.SwitchModel,
                    data.DailyClicks,
                    data.TotalClicks,
                    data.CreationDate.ToString("yyyy-MM-dd")
                );
            }

            if (dataGridView.Rows.Count > 0)
            {
                dataGridView.Rows[0].Selected = true;
                UpdateImagePath();
            }
        }

        private void DataGridView_SelectionChanged(object sender, EventArgs e)
        {
            UpdateImagePath();
        }

        private void UpdateImagePath()
        {
            if (dataGridView.SelectedRows.Count > 0)
            {
                string mouseModel = dataGridView.SelectedRows[0].Cells["MouseModel"].Value.ToString();
                if (mouseDataDict.ContainsKey(mouseModel))
                {
                    txtImagePath.Text = mouseDataDict[mouseModel].ImagePath ?? "";
                }
            }
        }

        private void BtnBrowseImage_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "圖片檔案|*.jpg;*.jpeg;*.png;*.gif;*.bmp";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                txtImagePath.Text = dialog.FileName;
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            try
            {
                foreach (DataGridViewRow row in dataGridView.Rows)
                {
                    string mouseModel = row.Cells["MouseModel"].Value.ToString();
                    if (mouseDataDict.ContainsKey(mouseModel))
                    {
                        mouseDataDict[mouseModel].SwitchBrand = row.Cells["SwitchBrand"].Value != null ? row.Cells["SwitchBrand"].Value.ToString() : "";
                        mouseDataDict[mouseModel].SwitchModel = row.Cells["SwitchModel"].Value != null ? row.Cells["SwitchModel"].Value.ToString() : "";
                    }
                }

                if (dataGridView.SelectedRows.Count > 0)
                {
                    string selectedMouseModel = dataGridView.SelectedRows[0].Cells["MouseModel"].Value.ToString();
                    if (mouseDataDict.ContainsKey(selectedMouseModel))
                    {
                        mouseDataDict[selectedMouseModel].ImagePath = txtImagePath.Text;
                    }
                }

                if (DataUpdated != null) DataUpdated(this, EventArgs.Empty);

                MessageBox.Show("設定已保存！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("保存失敗: " + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnReset_Click(object sender, EventArgs e)
        {
            if (dataGridView.SelectedRows.Count > 0)
            {
                DialogResult result = MessageBox.Show(
                    "確定要重置選中滑鼠的點擊記錄嗎？\n這將清除所有點擊數據！",
                    "確認重置", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                if (result == DialogResult.Yes)
                {
                    string mouseModel = dataGridView.SelectedRows[0].Cells["MouseModel"].Value.ToString();
                    if (mouseDataDict.ContainsKey(mouseModel))
                    {
                        mouseDataDict[mouseModel].DailyClicks = 0;
                        mouseDataDict[mouseModel].TotalClicks = 0;
                        mouseDataDict[mouseModel].CreationDate = DateTime.Now;
                        mouseDataDict[mouseModel].LastClickDate = DateTime.Now.Date;

                        LoadData();
                        if (DataUpdated != null) DataUpdated(this, EventArgs.Empty);
                        MessageBox.Show("記錄已重置！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            else
            {
                MessageBox.Show("請先選擇要重置的滑鼠記錄！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}