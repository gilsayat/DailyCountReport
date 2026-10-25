using System;
using System.IO;
using System.Drawing;
using System.Linq;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace DailyCountReport
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }

    public class AppSettings
    {
        public string LastExcelPath { get; set; } = "";
        public string LastTallyPath { get; set; } = "";
        public string DefaultSearchFolder { get; set; } = @"Z:\CC";

        private static string ConfigPath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "settings.json");

        public static AppSettings Load()
        {
            try
            {
                if (File.Exists(ConfigPath))
                {
                    string json = File.ReadAllText(ConfigPath);
                    return System.Text.Json.JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();
                }
            }
            catch { }
            return new AppSettings();
        }

        public void Save()
        {
            try
            {
                string json = System.Text.Json.JsonSerializer.Serialize(this, new System.Text.Json.JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(ConfigPath, json);
            }
            catch { }
        }
    }

    public class TallyItem
    {
        public string PhotoSlot { get; set; } = ""; // e.g. P01
        public int FileCount { get; set; } = 0;
        public string RawString { get; set; } = ""; // e.g. DAY_01,1C,P1,111001251,DuttonRanch
        public string PrefixCode { get; set; } = ""; // e.g. P1, DK1, EMB3, DR6
        public bool IsRightSide8x10 => PrefixCode.StartsWith("P", StringComparison.OrdinalIgnoreCase);
        public bool Is5x7 => PrefixCode.StartsWith("EMB", StringComparison.OrdinalIgnoreCase) ||
                             PrefixCode.StartsWith("CS", StringComparison.OrdinalIgnoreCase) ||
                             PrefixCode.StartsWith("DK", StringComparison.OrdinalIgnoreCase);
        public bool Is6x9 => PrefixCode.StartsWith("GW", StringComparison.OrdinalIgnoreCase) ||
                             PrefixCode.StartsWith("DR", StringComparison.OrdinalIgnoreCase);
    }

    public class MainForm : Form
    {
        private AppSettings settings = AppSettings.Load();

        private ComboBox cmbDay = null!;
        private TextBox txtExcelPath = null!;
        private TextBox txtTallyPath = null!;
        private Button btnBrowseExcel = null!;
        private Button btnBrowseTally = null!;
        private Button btnGenerate = null!;
        private Button btnSettings = null!;
        private ListBox lstLog = null!;

        private readonly Color darkBg = Color.FromArgb(32, 33, 36);
        private readonly Color panelBg = Color.FromArgb(44, 45, 49);
        private readonly Color accentBlue = Color.FromArgb(28, 115, 189);
        private readonly Color accentGreen = Color.FromArgb(34, 139, 34);
        private readonly Color textWhite = Color.White;

        public MainForm()
        {
            InitializeCustomComponent();
        }

        private void InitializeCustomComponent()
        {
            this.Text = "Daily Count Report Generator";
            this.Size = new Size(580, 480);
            this.MinimumSize = new Size(540, 440);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = darkBg;
            this.ForeColor = textWhite;
            this.Font = new Font("Segoe UI", 9.5f);

            Label lblDay = new Label { Text = "Select Report Day:", Location = new Point(25, 20), AutoSize = true, Font = new Font("Segoe UI", 10f, FontStyle.Bold) };
            cmbDay = new ComboBox { Location = new Point(170, 18), Width = 150, DropDownStyle = ComboBoxStyle.DropDownList };
            for (int i = 1; i <= 8; i++) cmbDay.Items.Add($"DAY {i}");
            cmbDay.SelectedIndex = 0;

            btnSettings = new Button { Text = "⚙ Settings", Location = new Point(440, 15), Width = 95, Height = 30, BackColor = panelBg, ForeColor = textWhite, FlatStyle = FlatStyle.Flat };
            btnSettings.Click += (s, e) => OpenSettingsDialog();

            Label lblExcel = new Label { Text = "Excel Master File (.xlsx):", Location = new Point(25, 65), AutoSize = true, Font = new Font("Segoe UI", 9f, FontStyle.Bold) };
            txtExcelPath = new TextBox { Location = new Point(25, 90), Width = 405, Text = settings.LastExcelPath };
            btnBrowseExcel = new Button { Text = "Browse...", Location = new Point(440, 88), Width = 95, Height = 28, BackColor = panelBg, ForeColor = textWhite, FlatStyle = FlatStyle.Flat };
            btnBrowseExcel.Click += (s, e) => BrowseExcelFile();

            Label lblTally = new Label { Text = "Tally Data File (.txt or .ps1):", Location = new Point(25, 130), AutoSize = true, Font = new Font("Segoe UI", 9f, FontStyle.Bold) };
            txtTallyPath = new TextBox { Location = new Point(25, 155), Width = 405, Text = settings.LastTallyPath };
            btnBrowseTally = new Button { Text = "Browse...", Location = new Point(440, 153), Width = 95, Height = 28, BackColor = panelBg, ForeColor = textWhite, FlatStyle = FlatStyle.Flat };
            btnBrowseTally.Click += (s, e) => BrowseTallyFile();

            btnGenerate = new Button { Text = "Generate & Update Excel Report", Location = new Point(25, 200), Width = 510, Height = 42, BackColor = accentGreen, ForeColor = textWhite, Font = new Font("Segoe UI", 11f, FontStyle.Bold), FlatStyle = FlatStyle.Flat };
            btnGenerate.Click += (s, e) => RunReportGeneration();

            Label lblLog = new Label { Text = "Process Log:", Location = new Point(25, 255), AutoSize = true, ForeColor = Color.LightGray };
            lstLog = new ListBox { Location = new Point(25, 280), Width = 510, Height = 140, BackColor = panelBg, ForeColor = Color.LightGreen, Font = new Font("Consolas", 8.5f) };

            this.Controls.Add(lblDay);
            this.Controls.Add(cmbDay);
            this.Controls.Add(btnSettings);
            this.Controls.Add(lblExcel);
            this.Controls.Add(txtExcelPath);
            this.Controls.Add(btnBrowseExcel);
            this.Controls.Add(lblTally);
            this.Controls.Add(txtTallyPath);
            this.Controls.Add(btnBrowseTally);
            this.Controls.Add(btnGenerate);
            this.Controls.Add(lblLog);
            this.Controls.Add(lstLog);
        }

        private void Log(string msg)
        {
            lstLog.Items.Add($"[{DateTime.Now:HH:mm:ss}] {msg}");
            lstLog.SelectedIndex = lstLog.Items.Count - 1;
        }

        private void BrowseExcelFile()
        {
            using var ofd = new OpenFileDialog
            {
                Title = "Select Daily Count Excel File",
                Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
            };
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                txtExcelPath.Text = ofd.FileName;
                settings.LastExcelPath = ofd.FileName;
                settings.Save();
            }
        }

        private void BrowseTallyFile()
        {
            using var ofd = new OpenFileDialog
            {
                Title = "Select Tally Text or PS1 File",
                Filter = "Tally Files (*.txt;*.ps1)|*.txt;*.ps1|All Files (*.*)|*.*"
            };
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                txtTallyPath.Text = ofd.FileName;
                settings.LastTallyPath = ofd.FileName;
                settings.Save();
            }
        }

        private void OpenSettingsDialog()
        {
            using Form dlg = new Form
            {
                Text = "Report Settings",
                Size = new Size(500, 220),
                StartPosition = FormStartPosition.CenterParent,
                BackColor = darkBg,
                ForeColor = textWhite,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false
            };

            Label lbl = new Label { Text = "Default Data Directory (Z:\\CC):", Location = new Point(20, 20), AutoSize = true };
            TextBox txt = new TextBox { Text = settings.DefaultSearchFolder, Location = new Point(20, 50), Width = 350 };
            Button btnB = new Button { Text = "Browse", Location = new Point(380, 48), Width = 80, Height = 26, BackColor = panelBg, ForeColor = textWhite, FlatStyle = FlatStyle.Flat };
            btnB.Click += (s, e) =>
            {
                using var fbd = new FolderBrowserDialog();
                if (fbd.ShowDialog() == DialogResult.OK) txt.Text = fbd.SelectedPath;
            };

            Button btnSave = new Button { Text = "Save Settings", Location = new Point(180, 110), Width = 130, Height = 35, BackColor = accentBlue, ForeColor = textWhite, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9.5f, FontStyle.Bold) };
            btnSave.Click += (s, e) =>
            {
                settings.DefaultSearchFolder = txt.Text;
                settings.Save();
                dlg.Close();
            };

            dlg.Controls.Add(lbl);
            dlg.Controls.Add(txt);
            dlg.Controls.Add(btnB);
            dlg.Controls.Add(btnSave);
            dlg.ShowDialog();
        }

        private void RunReportGeneration()
        {
            string excelPath = txtExcelPath.Text.Trim();
            string tallyPath = txtTallyPath.Text.Trim();
            string targetDay = cmbDay.SelectedItem?.ToString() ?? "DAY 1";

            if (!File.Exists(excelPath))
            {
                MessageBox.Show("Please select a valid Excel file.", "File Not Found", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!File.Exists(tallyPath))
            {
                MessageBox.Show("Please select a valid Tally (.txt) or PS1 file.", "File Not Found", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                Log($"Starting report update for sheet [{targetDay}]...");
                var items = ParseTallyFile(tallyPath);
                Log($"Parsed {items.Count} items from data source.");

                using var workbook = new XLWorkbook(excelPath);
                if (!workbook.Worksheets.TryGetWorksheet(targetDay, out var ws))
                {
                    MessageBox.Show($"Worksheet '{targetDay}' was not found in the workbook.", "Missing Sheet", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 1. Build row lookup for P01-P18 in the main block (rows 4 to 18)
                var slotRowMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                for (int r = 4; r <= 18; r++)
                {
                    string slotVal = ws.Cell(r, 7).GetString().Trim(); // Column G (PHOTO NO)
                    if (!string.IsNullOrEmpty(slotVal))
                    {
                        slotRowMap[slotVal] = r;
                    }
                }

                int sum5x7 = 0;
                int rightCount = 0;
                int leftCount = 0;
                int nextSecondaryRow = 20; // Secondary left assignments start at row 20

                foreach (var item in items)
                {
                    if (item.IsRightSide8x10)
                    {
                        // Right side: Column O (15) = Count, Column P (16) = RawString
                        if (slotRowMap.TryGetValue(item.PhotoSlot, out int targetRow))
                        {
                            ws.Cell(targetRow, 15).Value = item.FileCount;
                            ws.Cell(targetRow, 16).Value = item.RawString;
                            rightCount++;
                        }
                    }
                    else
                    {
                        // Left side (5x7 & 6x9): Column H (8) = Count, Column I (9) = RawString
                        if (item.Is5x7)
                        {
                            sum5x7 += item.FileCount;
                        }

                        if (slotRowMap.TryGetValue(item.PhotoSlot, out int targetRow))
                        {
                            // If primary slot is empty, put it there
                            if (ws.Cell(targetRow, 8).IsEmpty() && ws.Cell(targetRow, 9).IsEmpty())
                            {
                                ws.Cell(targetRow, 8).Value = item.FileCount;
                                ws.Cell(targetRow, 9).Value = item.RawString;
                                leftCount++;
                            }
                            else
                            {
                                // Main slot already used; place in secondary area (rows 20-25)
                                if (nextSecondaryRow <= 25)
                                {
                                    ws.Cell(nextSecondaryRow, 5).Value = ws.Cell(targetRow, 5).Value; // Copy Photographer Name
                                    ws.Cell(nextSecondaryRow, 6).Value = ws.Cell(targetRow, 6).Value; // Copy Crew ID
                                    ws.Cell(nextSecondaryRow, 7).Value = item.PhotoSlot;
                                    ws.Cell(nextSecondaryRow, 8).Value = item.FileCount;
                                    ws.Cell(nextSecondaryRow, 9).Value = item.RawString;
                                    nextSecondaryRow++;
                                    leftCount++;
                                }
                            }
                        }
                    }
                }

                // Write the 5x7 sum directly into Cell B11
                ws.Cell(11, 2).Value = sum5x7;

                workbook.Save();
                Log($"Success: {rightCount} 8x10 items placed on Right side.");
                Log($"Success: {leftCount} 5x7/6x9 items placed on Left side.");
                Log($"5x7 Total (Cell B11) set to: {sum5x7}");
                Log($"Worksheet [{targetDay}] updated and saved successfully!");

                MessageBox.Show($"Report for {targetDay} generated successfully!\n\n5x7 Total: {sum5x7}\n8x10 Items: {rightCount}\n5x7/6x9 Items: {leftCount}", "Report Ready", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Log($"Error: {ex.Message}");
                MessageBox.Show($"Failed to generate report: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private List<TallyItem> ParseTallyFile(string filePath)
        {
            var list = new List<TallyItem>();
            string[] lines = File.ReadAllLines(filePath);

            foreach (var line in lines)
            {
                string tr = line.Trim();
                if (string.IsNullOrWhiteSpace(tr) || tr.StartsWith("*") || tr.StartsWith("#")) continue;

                // Tab or space-delimited tally format: P01 \t 36 \t DAY_02,2E,DK1,14115,WaterWorks
                var parts = tr.Split(new[] { '\t' }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length >= 3)
                {
                    string slot = parts[0].Trim();
                    if (!int.TryParse(parts[1].Trim(), out int count)) count = 0;
                    string rawStr = parts[2].Trim();

                    string prefix = ExtractPrefixCode(rawStr);
                    list.Add(new TallyItem { PhotoSlot = slot, FileCount = count, RawString = rawStr, PrefixCode = prefix });
                }
                else if (tr.Contains("PathPhoto") && tr.Contains("="))
                {
                    // PS1 format: $PathPhoto01 = "DAY_02,2E,DK1,14115,WaterWorks"
                    var m = Regex.Match(tr, @"\$PathPhoto(\d+)\s*=\s*""([^""]+)""");
                    if (m.Success)
                    {
                        string slot = "P" + m.Groups[1].Value;
                        string rawStr = m.Groups[2].Value.Trim();
                        string prefix = ExtractPrefixCode(rawStr);
                        list.Add(new TallyItem { PhotoSlot = slot, FileCount = 0, RawString = rawStr, PrefixCode = prefix });
                    }
                }
            }
            return list;
        }

        private string ExtractPrefixCode(string rawStr)
        {
            if (string.IsNullOrEmpty(rawStr)) return "";
            var subParts = rawStr.Split(',');
            return subParts.Length >= 3 ? subParts[2].Trim() : "";
        }
    }
}
