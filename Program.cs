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
        public string PhotoSlot { get; set; } = "";
        public int FileCount { get; set; } = 0;
        public string RawString { get; set; } = "";
        public string PrefixCode { get; set; } = "";
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
            this.Size = new Size(540, 470);
            this.MinimumSize = new Size(500, 430);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = darkBg;
            this.ForeColor = textWhite;
            this.Font = new Font("Segoe UI", 9.5f);

            Label lblDay = new Label { Text = "Select Day Sheet:", Location = new Point(20, 20), AutoSize = true, Font = new Font("Segoe UI", 10f, FontStyle.Bold) };
            cmbDay = new ComboBox { Location = new Point(160, 18), Width = 150, DropDownStyle = ComboBoxStyle.DropDownList };
            for (int i = 1; i <= 8; i++) cmbDay.Items.Add($"DAY {i}");
            cmbDay.SelectedIndex = 0;

            Label lblExcel = new Label { Text = "Excel Master File (.xlsx):", Location = new Point(20, 65), AutoSize = true, Font = new Font("Segoe UI", 9f, FontStyle.Bold) };
            txtExcelPath = new TextBox { Location = new Point(20, 90), Width = 375, Text = settings.LastExcelPath };
            btnBrowseExcel = new Button { Text = "Browse...", Location = new Point(405, 88), Width = 95, Height = 28, BackColor = panelBg, ForeColor = textWhite, FlatStyle = FlatStyle.Flat };
            btnBrowseExcel.Click += (s, e) => BrowseExcelFile();

            Label lblTally = new Label { Text = "Tally Data File (.txt or .ps1):", Location = new Point(20, 130), AutoSize = true, Font = new Font("Segoe UI", 9f, FontStyle.Bold) };
            txtTallyPath = new TextBox { Location = new Point(20, 155), Width = 375, Text = settings.LastTallyPath };
            btnBrowseTally = new Button { Text = "Browse...", Location = new Point(405, 153), Width = 95, Height = 28, BackColor = panelBg, ForeColor = textWhite, FlatStyle = FlatStyle.Flat };
            btnBrowseTally.Click += (s, e) => BrowseTallyFile();

            btnGenerate = new Button { Text = "Generate & Update Excel Report", Location = new Point(20, 200), Width = 480, Height = 42, BackColor = accentGreen, ForeColor = textWhite, Font = new Font("Segoe UI", 11f, FontStyle.Bold), FlatStyle = FlatStyle.Flat };
            btnGenerate.Click += (s, e) => RunReportGeneration();

            Label lblLog = new Label { Text = "Status Log:", Location = new Point(20, 255), AutoSize = true, ForeColor = Color.LightGray };
            lstLog = new ListBox { Location = new Point(20, 280), Width = 480, Height = 130, BackColor = panelBg, ForeColor = Color.LightGreen, Font = new Font("Consolas", 8.5f) };

            this.Controls.Add(lblDay);
            this.Controls.Add(cmbDay);
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
            using var ofd = new OpenFileDialog { Title = "Select Daily Count Excel File", Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*" };
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                txtExcelPath.Text = ofd.FileName;
                settings.LastExcelPath = ofd.FileName;
                settings.Save();
            }
        }

        private void BrowseTallyFile()
        {
            using var ofd = new OpenFileDialog { Title = "Select Tally Text or PS1 File", Filter = "Tally Files (*.txt;*.ps1)|*.txt;*.ps1|All Files (*.*)|*.*" };
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                txtTallyPath.Text = ofd.FileName;
                settings.LastTallyPath = ofd.FileName;
                settings.Save();
            }
        }

        private void SetBoldCenter(IXLCell cell, object val)
        {
            cell.Value = XLCellValue.FromObject(val);
            cell.Style.Font.SetBold(true);
            cell.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            cell.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
        }

        private void RunReportGeneration()
        {
            string excelPath = txtExcelPath.Text.Trim();
            string tallyPath = txtTallyPath.Text.Trim();
            string targetDay = cmbDay.SelectedItem?.ToString() ?? "DAY 1";

            if (!File.Exists(excelPath))
            {
                MessageBox.Show("Please select an existing Excel file.", "Missing File", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!File.Exists(tallyPath))
            {
                MessageBox.Show("Please select an existing Tally or PS1 file.", "Missing File", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                Log($"Reading data from [{Path.GetFileName(tallyPath)}]...");
                var items = ParseTallyFile(tallyPath);
                Log($"Found {items.Count} photo assignments.");

                using var workbook = new XLWorkbook(excelPath);
                if (!workbook.Worksheets.TryGetWorksheet(targetDay, out var ws))
                {
                    MessageBox.Show($"Worksheet '{targetDay}' does not exist in this workbook.", "Sheet Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Map photo numbers (P01-P18) to rows in the main block
                var slotRowMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                for (int r = 4; r <= 18; r++)
                {
                    string slotVal = ws.Cell(r, 7).GetString().Trim();
                    if (!string.IsNullOrEmpty(slotVal)) slotRowMap[slotVal] = r;
                }

                // Find the TOTAL row (looks for "TOTAL" in Col E)
                int totalRow = 24;
                for (int r = 18; r <= 35; r++)
                {
                    if (ws.Cell(r, 5).GetString().Trim().Equals("TOTAL", StringComparison.OrdinalIgnoreCase))
                    {
                        totalRow = r;
                        break;
                    }
                }

                int totalLeftCount = 0;
                int totalRightCount = 0;
                int sum5x7 = 0;
                int sum6x9 = 0;
                int nextSecondaryRow = 19;

                foreach (var item in items)
                {
                    if (item.IsRightSide8x10)
                    {
                        // Right side: Col N (14) = Session (Count / 8), Col O (15) = Count, Col P (16) = RawString
                        if (slotRowMap.TryGetValue(item.PhotoSlot, out int targetRow))
                        {
                            int sessionVal = (int)Math.Round((double)item.FileCount / 8.0);
                            SetBoldCenter(ws.Cell(targetRow, 14), sessionVal);
                            SetBoldCenter(ws.Cell(targetRow, 15), item.FileCount);
                            SetBoldCenter(ws.Cell(targetRow, 16), item.RawString);
                            totalRightCount += item.FileCount;
                        }
                    }
                    else
                    {
                        // Left side (5x7 & 6x9)
                        if (item.Is5x7) sum5x7 += item.FileCount;
                        if (item.Is6x9) sum6x9 += item.FileCount;
                        totalLeftCount += item.FileCount;

                        if (slotRowMap.TryGetValue(item.PhotoSlot, out int targetRow))
                        {
                            // If primary slot is empty, fill it
                            if (ws.Cell(targetRow, 8).IsEmpty() && ws.Cell(targetRow, 9).IsEmpty())
                            {
                                SetBoldCenter(ws.Cell(targetRow, 8), item.FileCount);
                                SetBoldCenter(ws.Cell(targetRow, 9), item.RawString);
                            }
                            else
                            {
                                // Otherwise place in secondary rows (below row 18, above TOTAL)
                                if (nextSecondaryRow < totalRow)
                                {
                                    SetBoldCenter(ws.Cell(nextSecondaryRow, 5), ws.Cell(targetRow, 5).Value);
                                    SetBoldCenter(ws.Cell(nextSecondaryRow, 6), ws.Cell(targetRow, 6).Value);
                                    SetBoldCenter(ws.Cell(nextSecondaryRow, 7), item.PhotoSlot);
                                    SetBoldCenter(ws.Cell(nextSecondaryRow, 8), item.FileCount);
                                    SetBoldCenter(ws.Cell(nextSecondaryRow, 9), item.RawString);
                                    nextSecondaryRow++;
                                }
                            }
                        }
                    }
                }

                // Put total file counts directly in the TOTAL row
                SetBoldCenter(ws.Cell(totalRow, 8), totalLeftCount);
                SetBoldCenter(ws.Cell(totalRow, 15), totalRightCount);

                // Update side summary totals in Column B
                SetBoldCenter(ws.Cell(11, 2), sum5x7);                         // 5x7
                SetBoldCenter(ws.Cell(13, 2), sum6x9);                         // 6x9
                SetBoldCenter(ws.Cell(15, 2), totalRightCount);                // 8x10
                SetBoldCenter(ws.Cell(7, 2), totalLeftCount + totalRightCount); // Grand Total

                // Configure Page Setup for Printing (Both sections Landscape, Narrow Margins, Fit to 1 Page Wide)
                var pageSetup = ws.PageSetup;
                pageSetup.PageOrientation = XLPageOrientation.Landscape;
                pageSetup.PagesWide = 1;
                pageSetup.PagesTall = 0; // Automatic height
                pageSetup.Margins.Top = 0.25;
                pageSetup.Margins.Bottom = 0.25;
                pageSetup.Margins.Left = 0.25;
                pageSetup.Margins.Right = 0.25;

                workbook.Save();

                Log($"Updated {targetDay} successfully!");
                Log($"Left Total (Col H): {totalLeftCount} | Right Total (Col O): {totalRightCount}");
                Log($"5x7: {sum5x7} | 6x9: {sum6x9} | Grand Total: {totalLeftCount + totalRightCount}");

                MessageBox.Show($"Report for {targetDay} generated!\n\n" +
                                $"Left File Count: {totalLeftCount}\n" +
                                $"Right File Count: {totalRightCount}\n" +
                                $"5x7 Total: {sum5x7}\n" +
                                $"6x9 Total: {sum6x9}\n" +
                                $"Grand Total: {totalLeftCount + totalRightCount}",
                                "Report Updated", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Log($"Error: {ex.Message}");
                MessageBox.Show($"Error generating report: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
