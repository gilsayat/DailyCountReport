using System;
using System.IO;
using System.Drawing;
using System.Linq;
using System.Diagnostics;
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
        public bool IsNightEvent { get; set; } = false;

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
        private Button btnOpenExcel = null!;
        private Label lblVoyageStatus = null!;
        private ListBox lstLog = null!;

        private readonly Color darkBg = Color.FromArgb(32, 33, 36);
        private readonly Color panelBg = Color.FromArgb(44, 45, 49);
        private readonly Color accentBlue = Color.FromArgb(28, 115, 189);
        private readonly Color accentGreen = Color.FromArgb(34, 139, 34);
        private readonly Color textWhite = Color.White;

        public MainForm()
        {
            InitializeCustomComponent();
            SetupDragAndDrop();
            ValidateVoyageMatch();
        }

        private void InitializeCustomComponent()
        {
            this.Text = "Daily Count Report Generator";
            this.Size = new Size(580, 520);
            this.MinimumSize = new Size(540, 480);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = darkBg;
            this.ForeColor = textWhite;
            this.Font = new Font("Segoe UI", 9.5f);

            Label lblDay = new Label { Text = "Report Day Sheet:", Location = new Point(20, 20), AutoSize = true, Font = new Font("Segoe UI", 10f, FontStyle.Bold) };
            cmbDay = new ComboBox { Location = new Point(160, 18), Width = 140, DropDownStyle = ComboBoxStyle.DropDownList };
            for (int i = 1; i <= 8; i++) cmbDay.Items.Add($"DAY {i}");
            cmbDay.SelectedIndex = 0;

            lblVoyageStatus = new Label { Text = "Voyage: Ready", Location = new Point(315, 21), AutoSize = true, ForeColor = Color.LightGray, Font = new Font("Segoe UI", 9f, FontStyle.Italic) };

            Label lblExcel = new Label { Text = "Excel Master File (.xlsx) - [Drag & Drop]:", Location = new Point(20, 65), AutoSize = true, Font = new Font("Segoe UI", 9f, FontStyle.Bold) };
            txtExcelPath = new TextBox { Location = new Point(20, 90), Width = 415, Text = settings.LastExcelPath };
            btnBrowseExcel = new Button { Text = "Browse...", Location = new Point(445, 88), Width = 95, Height = 28, BackColor = panelBg, ForeColor = textWhite, FlatStyle = FlatStyle.Flat };
            btnBrowseExcel.Click += (s, e) => BrowseExcelFile();

            Label lblTally = new Label { Text = "Tally Data File (.txt or .ps1) - [Drag & Drop]:", Location = new Point(20, 130), AutoSize = true, Font = new Font("Segoe UI", 9f, FontStyle.Bold) };
            txtTallyPath = new TextBox { Location = new Point(20, 155), Width = 415, Text = settings.LastTallyPath };
            btnBrowseTally = new Button { Text = "Browse...", Location = new Point(445, 153), Width = 95, Height = 28, BackColor = panelBg, ForeColor = textWhite, FlatStyle = FlatStyle.Flat };
            btnBrowseTally.Click += (s, e) => BrowseTallyFile();

            btnGenerate = new Button { Text = "⚡ Generate & Update Excel Report", Location = new Point(20, 200), Width = 385, Height = 42, BackColor = accentGreen, ForeColor = textWhite, Font = new Font("Segoe UI", 11f, FontStyle.Bold), FlatStyle = FlatStyle.Flat };
            btnGenerate.Click += (s, e) => RunReportGeneration();

            btnOpenExcel = new Button { Text = "Open File", Location = new Point(415, 200), Width = 125, Height = 42, BackColor = accentBlue, ForeColor = textWhite, Font = new Font("Segoe UI", 9.5f, FontStyle.Bold), FlatStyle = FlatStyle.Flat, Enabled = File.Exists(settings.LastExcelPath) };
            btnOpenExcel.Click += (s, e) => OpenExcelWorkbook();

            Label lblLog = new Label { Text = "Status Log:", Location = new Point(20, 255), AutoSize = true, ForeColor = Color.LightGray };
            lstLog = new ListBox { Location = new Point(20, 280), Width = 520, Height = 175, BackColor = panelBg, ForeColor = Color.LightGreen, Font = new Font("Consolas", 8.5f) };

            this.Controls.Add(lblDay);
            this.Controls.Add(cmbDay);
            this.Controls.Add(lblVoyageStatus);
            this.Controls.Add(lblExcel);
            this.Controls.Add(txtExcelPath);
            this.Controls.Add(btnBrowseExcel);
            this.Controls.Add(lblTally);
            this.Controls.Add(txtTallyPath);
            this.Controls.Add(btnBrowseTally);
            this.Controls.Add(btnGenerate);
            this.Controls.Add(btnOpenExcel);
            this.Controls.Add(lblLog);
            this.Controls.Add(lstLog);
        }

        private void SetupDragAndDrop()
        {
            this.AllowDrop = true;
            this.DragEnter += (s, e) =>
            {
                if (e.Data != null && e.Data.GetDataPresent(DataFormats.FileDrop))
                    e.Effect = DragDropEffects.Copy;
            };

            this.DragDrop += (s, e) =>
            {
                if (e.Data != null && e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    string[] files = (string[])e.Data.GetData(DataFormats.FileDrop)!;
                    foreach (var file in files)
                    {
                        string ext = Path.GetExtension(file).ToLower();
                        if (ext == ".xlsx" || ext == ".xlsm")
                        {
                            txtExcelPath.Text = file;
                            settings.LastExcelPath = file;
                            settings.Save();
                            btnOpenExcel.Enabled = true;
                            Log($"Loaded Excel workbook: {Path.GetFileName(file)}");
                            ValidateVoyageMatch();
                        }
                        else if (ext == ".txt" || ext == ".ps1")
                        {
                            ProcessTallySelection(file);
                        }
                    }
                }
            };
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
                btnOpenExcel.Enabled = true;
                Log($"Selected Excel: {Path.GetFileName(ofd.FileName)}");
                ValidateVoyageMatch();
            }
        }

        private void BrowseTallyFile()
        {
            using var ofd = new OpenFileDialog { Title = "Select Tally Text or PS1 File", Filter = "Tally Files (*.txt;*.ps1)|*.txt;*.ps1|All Files (*.*)|*.*" };
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                ProcessTallySelection(ofd.FileName);
            }
        }

        private void ProcessTallySelection(string filePath)
        {
            txtTallyPath.Text = filePath;
            settings.LastTallyPath = filePath;
            settings.Save();
            Log($"Loaded Tally: {Path.GetFileName(filePath)}");

            int detectedDay = ExtractDayNumber(filePath);
            if (detectedDay >= 1 && detectedDay <= 8)
            {
                cmbDay.SelectedItem = $"DAY {detectedDay}";
                Log($"Auto-detected target sheet: [DAY {detectedDay}]");
            }

            if (string.IsNullOrWhiteSpace(txtExcelPath.Text) || !File.Exists(txtExcelPath.Text))
            {
                string dir = Path.GetDirectoryName(filePath) ?? "";
                if (Directory.Exists(dir))
                {
                    var found = Directory.GetFiles(dir, "*Daily Count*.xlsx").FirstOrDefault()
                             ?? Directory.GetFiles(dir, "*.xlsx").FirstOrDefault();
                    if (!string.IsNullOrEmpty(found))
                    {
                        txtExcelPath.Text = found;
                        settings.LastExcelPath = found;
                        settings.Save();
                        btnOpenExcel.Enabled = true;
                        Log($"Auto-detected matching Excel file: {Path.GetFileName(found)}");
                    }
                }
            }
            ValidateVoyageMatch();
        }

        private void ValidateVoyageMatch()
        {
            try
            {
                if (File.Exists(txtTallyPath.Text) && File.Exists(txtExcelPath.Text))
                {
                    string tallyName = Path.GetFileName(txtTallyPath.Text);
                    string excelName = Path.GetFileName(txtExcelPath.Text);

                    var m1 = Regex.Match(tallyName, @"([A-Z]{2}\d{6})", RegexOptions.IgnoreCase);
                    var m2 = Regex.Match(excelName, @"([A-Z]{2}\d{6})", RegexOptions.IgnoreCase);

                    if (m1.Success && m2.Success && m1.Value.Equals(m2.Value, StringComparison.OrdinalIgnoreCase))
                    {
                        lblVoyageStatus.Text = $"Voyage: {m1.Value.ToUpper()} (Matched ✓)";
                        lblVoyageStatus.ForeColor = Color.FromArgb(120, 240, 150);
                        return;
                    }
                }
            }
            catch { }
            lblVoyageStatus.Text = "Voyage: Ready";
            lblVoyageStatus.ForeColor = Color.LightGray;
        }

        private void OpenExcelWorkbook()
        {
            try
            {
                if (File.Exists(txtExcelPath.Text))
                {
                    Process.Start(new ProcessStartInfo { FileName = txtExcelPath.Text, UseShellExecute = true });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Could not open file: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private int ExtractDayNumber(string path)
        {
            string fileName = Path.GetFileNameWithoutExtension(path);
            var m = Regex.Match(fileName, @"(?:DAY[_\s]*|_0?)([1-8])\b", RegexOptions.IgnoreCase);
            if (m.Success && int.TryParse(m.Groups[1].Value, out int day))
            {
                return day;
            }
            return 0;
        }

        private void SetStyledCell(IXLCell cell, object val, XLAlignmentHorizontalValues hAlign)
        {
            cell.Value = XLCellValue.FromObject(val);
            cell.Style.Font.SetBold(true);
            cell.Style.Alignment.SetHorizontal(hAlign);
            cell.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
        }

        private void RunReportGeneration()
        {
            string excelPath = txtExcelPath.Text.Trim();
            string tallyPath = txtTallyPath.Text.Trim();
            string selectedDay = cmbDay.SelectedItem?.ToString() ?? "DAY 1";

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

            // Safeguard 1: Day mismatch check
            int detectedDay = ExtractDayNumber(tallyPath);
            if (detectedDay >= 1 && detectedDay <= 8 && selectedDay != $"DAY {detectedDay}")
            {
                var dr = MessageBox.Show(
                    $"Day Mismatch Detected!\n\n" +
                    $"The tally file appears to be for [DAY {detectedDay}], but you selected [{selectedDay}].\n\n" +
                    $"Click 'Yes' to automatically switch to [DAY {detectedDay}].\n" +
                    $"Click 'No' to proceed with [{selectedDay}].\n" +
                    $"Click 'Cancel' to abort.",
                    "Day Mismatch Alert", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);

                if (dr == DialogResult.Cancel) return;
                if (dr == DialogResult.Yes)
                {
                    selectedDay = $"DAY {detectedDay}";
                    cmbDay.SelectedItem = selectedDay;
                }
            }

            try
            {
                using var workbook = new XLWorkbook(excelPath);
                if (!workbook.Worksheets.TryGetWorksheet(selectedDay, out var ws))
                {
                    MessageBox.Show($"Worksheet '{selectedDay}' does not exist in this workbook.", "Sheet Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Locate the dynamic TOTAL row
                int totalRow = 24;
                for (int r = 18; r <= 35; r++)
                {
                    if (ws.Cell(r, 5).GetString().Trim().Equals("TOTAL", StringComparison.OrdinalIgnoreCase))
                    {
                        totalRow = r;
                        break;
                    }
                }

                // Safeguard 2: Overwrite Protection check
                bool hasExistingData = !ws.Cell(4, 8).IsEmpty() || !ws.Cell(4, 15).IsEmpty();
                if (hasExistingData)
                {
                    var confirm = MessageBox.Show(
                        $"Worksheet '{selectedDay}' already contains data.\n\nDo you want to clear and overwrite it with the updated tally?",
                        "Confirm Overwrite", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (confirm != DialogResult.Yes) return;

                    Log($"Clearing previous report data on [{selectedDay}]...");

                    // 1. Clear data columns in primary block (rows 4 to 18)
                    for (int r = 4; r <= 18; r++)
                    {
                        ws.Cell(r, 8).Clear();   // Left count
                        ws.Cell(r, 9).Clear();   // Left string
                        ws.Cell(r, 14).Clear();  // Session
                        ws.Cell(r, 15).Clear();  // Right count
                        ws.Cell(r, 16).Clear();  // Right string
                    }

                    // 2. Completely clear secondary rows (row 19 up to totalRow - 1)
                    for (int r = 19; r < totalRow; r++)
                    {
                        for (int c = 5; c <= 16; c++) ws.Cell(r, c).Clear();
                    }

                    // 3. Reset total row and summary figures
                    ws.Cell(totalRow, 8).Clear();
                    ws.Cell(totalRow, 15).Clear();
                    ws.Cell(7, 2).Clear();
                    ws.Cell(11, 2).Clear();
                    ws.Cell(13, 2).Clear();
                    ws.Cell(15, 2).Clear();
                }

                Log($"Parsing assignments from [{Path.GetFileName(tallyPath)}]...");
                var items = ParseTallyFile(tallyPath);
                Log($"Found {items.Count} items.");

                // Map photo numbers (P01-P18) to rows in the primary block
                var slotRowMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                for (int r = 4; r <= 18; r++)
                {
                    string slotVal = ws.Cell(r, 7).GetString().Trim();
                    if (!string.IsNullOrEmpty(slotVal)) slotRowMap[slotVal] = r;
                }

                int totalLeftCount = 0;
                int totalRightCount = 0;
                int sum5x7 = 0;
                int sum6x9 = 0;

                // Row 19 stays completely blank as a clean separator.
                // Night event shifts on the left start at row 20.
                int nextSecondaryRow = 20;

                foreach (var item in items)
                {
                    if (item.FileCount == 0 && string.IsNullOrWhiteSpace(item.RawString))
                        continue;

                    if (item.IsRightSide8x10)
                    {
                        // Right side: Col N (14) = Session, Col O (15) = Count, Col P (16) = RawString
                        if (slotRowMap.TryGetValue(item.PhotoSlot, out int targetRow))
                        {
                            int sessionVal = (int)Math.Round((double)item.FileCount / 8.0);
                            SetStyledCell(ws.Cell(targetRow, 14), sessionVal, XLAlignmentHorizontalValues.Center);
                            SetStyledCell(ws.Cell(targetRow, 15), item.FileCount, XLAlignmentHorizontalValues.Right);
                            SetStyledCell(ws.Cell(targetRow, 16), item.RawString, XLAlignmentHorizontalValues.Left);
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
                            // Day Event items go into primary rows (4 to 18)
                            if (!item.IsNightEvent && ws.Cell(targetRow, 8).IsEmpty() && ws.Cell(targetRow, 9).IsEmpty())
                            {
                                SetStyledCell(ws.Cell(targetRow, 8), item.FileCount, XLAlignmentHorizontalValues.Right);
                                SetStyledCell(ws.Cell(targetRow, 9), item.RawString, XLAlignmentHorizontalValues.Left);
                            }
                            else
                            {
                                // Night Event items go to secondary rows starting from Row 20
                                if (nextSecondaryRow >= totalRow)
                                {
                                    ws.Row(totalRow).InsertRowsAbove(1);
                                    totalRow++;
                                }

                                SetStyledCell(ws.Cell(nextSecondaryRow, 5), ws.Cell(targetRow, 5).Value, XLAlignmentHorizontalValues.Center);
                                SetStyledCell(ws.Cell(nextSecondaryRow, 6), ws.Cell(targetRow, 6).Value, XLAlignmentHorizontalValues.Center);
                                SetStyledCell(ws.Cell(nextSecondaryRow, 7), item.PhotoSlot, XLAlignmentHorizontalValues.Center);
                                SetStyledCell(ws.Cell(nextSecondaryRow, 8), item.FileCount, XLAlignmentHorizontalValues.Right);
                                SetStyledCell(ws.Cell(nextSecondaryRow, 9), item.RawString, XLAlignmentHorizontalValues.Left);
                                nextSecondaryRow++;
                            }
                        }
                    }
                }

                // Bottom TOTAL row: Direct sum placement
                SetStyledCell(ws.Cell(totalRow, 8), totalLeftCount, XLAlignmentHorizontalValues.Right);
                SetStyledCell(ws.Cell(totalRow, 15), totalRightCount, XLAlignmentHorizontalValues.Right);

                // Side Summary Totals (Column B)
                SetStyledCell(ws.Cell(11, 2), sum5x7, XLAlignmentHorizontalValues.Right);
                SetStyledCell(ws.Cell(13, 2), sum6x9, XLAlignmentHorizontalValues.Right);
                SetStyledCell(ws.Cell(15, 2), totalRightCount, XLAlignmentHorizontalValues.Right);
                SetStyledCell(ws.Cell(7, 2), totalLeftCount + totalRightCount, XLAlignmentHorizontalValues.Right);

                // Print Page Setup: Landscape, Narrow Margins, Fit to 1 Page Wide
                var pageSetup = ws.PageSetup;
                pageSetup.PageOrientation = XLPageOrientation.Landscape;
                pageSetup.PagesWide = 1;
                pageSetup.PagesTall = 0;
                pageSetup.Margins.Top = 0.25;
                pageSetup.Margins.Bottom = 0.25;
                pageSetup.Margins.Left = 0.25;
                pageSetup.Margins.Right = 0.25;

                workbook.Save();

                Log($"✓ Updated [{selectedDay}] successfully!");
                Log($"Left Total: {totalLeftCount} | Right Total: {totalRightCount}");
                Log($"5x7: {sum5x7} | 6x9: {sum6x9} | Grand Total: {totalLeftCount + totalRightCount}");
                btnOpenExcel.Enabled = true;

                MessageBox.Show($"Report for {selectedDay} generated successfully!\n\n" +
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
            bool isNightEvent = false;

            foreach (var line in lines)
            {
                string tr = line.Trim();
                if (string.IsNullOrWhiteSpace(tr)) continue;

                if (tr.IndexOf("Night Event", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    isNightEvent = true;
                    continue;
                }
                if (tr.IndexOf("Day Event", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    isNightEvent = false;
                    continue;
                }

                if (tr.StartsWith("*") || tr.StartsWith("#")) continue;

                var parts = tr.Split(new[] { '\t' }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length >= 3)
                {
                    string slot = parts[0].Trim();
                    if (!int.TryParse(parts[1].Trim(), out int count)) count = 0;
                    string rawStr = parts[2].Trim();
                    string prefix = ExtractPrefixCode(rawStr);
                    list.Add(new TallyItem
                    {
                        PhotoSlot = slot,
                        FileCount = count,
                        RawString = rawStr,
                        PrefixCode = prefix,
                        IsNightEvent = isNightEvent
                    });
                }
                else if (tr.Contains("PathPhoto") && tr.Contains("="))
                {
                    var m = Regex.Match(tr, @"\$PathPhoto(\d+)\s*=\s*""([^""]+)""");
                    if (m.Success)
                    {
                        string slot = "P" + m.Groups[1].Value;
                        string rawStr = m.Groups[2].Value.Trim();
                        string prefix = ExtractPrefixCode(rawStr);
                        list.Add(new TallyItem
                        {
                            PhotoSlot = slot,
                            FileCount = 0,
                            RawString = rawStr,
                            PrefixCode = prefix,
                            IsNightEvent = isNightEvent
                        });
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
