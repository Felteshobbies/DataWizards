using libDataWizard;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace DataWizard.UI
{
    public partial class fMain : Form
    {

        private FileSystemWatcher fileWatcher;
        private Timer watcherTimer;
        private bool isWatchMode = false;
        private bool isOverwrite = false;

        private List<string> currentFiles = new List<string>();
        private string watchFolder = string.Empty;
        private string outputFolder = string.Empty;

        private Encoding optionCsvEncoding = null;
        private char optionCsvSeparator = '\0';
        private int optionCsvHeader = -1;

        private Encoding optionXlsEncoding = null;
        private char optionXlsSeparator = '\0';
        private bool optionXlsQuoteAllText = false;
        private bool optionXlsAllSheets = false;


        public fMain()
        {
            InitializeComponent();

            this.Text = "DataWizard - CSV ↔ Excel Converter";
            this.MinimumSize = new Size(800, 300);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.White;

            // Allow Drag & Drop
            this.AllowDrop = true;
            this.DragEnter += MainForm_DragEnter;
            this.DragDrop += MainForm_DragDrop;

            this.ResumeLayout(false);

            convertTab.Text = "Convert";
            settingsTab.Text = "Settings";
            watchTab.Text = "Watcher";
            logTab.Text = "Log";

            // Status Bar
            StatusStrip statusStrip = new StatusStrip();
            ToolStripStatusLabel statusLabel = new ToolStripStatusLabel("Ready");
            statusLabel.Name = "statusLabel";
            statusStrip.Items.Add(statusLabel);
            this.Controls.Add(statusStrip);

            dragDropPanel.DragEnter += DragDropPanel_DragEnter;
            dragDropPanel.DragDrop += DragDropPanel_DragDrop;
            dragDropPanel.Paint += DragDropPanel_Paint;

            convertButton.Enabled = false;

            cbCsvEncoding.Items.AddRange(new object[]
            {
                "Auto",
                "UTF-8",
                "ISO-8859-1",
                "Windows-1252"
             });
            cbCsvEncoding.SelectedIndex = 0;
            cbXlsEncoding.Items.AddRange(new object[]
            {
                "UTF-8",
                "UTF-8 mit BOM",
                "Windows-1252",
                "ISO-8859-1"
             });
            cbXlsEncoding.SelectedIndex = 0;
            cbCsvSeparator.Items.AddRange(new object[]
           {
                "Auto",
                "Semicolon (;)",
                "Comma (,)",
                "Tab",
                "Pipe (|)"
           });
            cbCsvSeparator.SelectedIndex = 0;
            cbXlsSeparator.Items.AddRange(new object[]
           {
                "Semicolon (;)",
                "Comma (,)",
                "Tab",
                "Pipe (|)"
           });
            cbXlsSeparator.SelectedIndex = 0;

            cbHeader.Items.AddRange(new object[]
             {
                "Auto",
                "Header",
                "No header"

              });
            cbHeader.SelectedIndex = 0;

            AddLog("DataWizard ready.");
        }

        private void DragDropPanel_Paint(object sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.Clear(Color.WhiteSmoke);

            // Gestrichelter Rahmen
            using (Pen dashedPen = new Pen(Color.LightGreen, 4))
            {
                dashedPen.DashStyle = System.Drawing.Drawing2D.DashStyle.Dash;
                g.DrawRectangle(dashedPen, 5, 5, dragDropPanel.Width - 10, dragDropPanel.Height - 10);
            }

            // Icon und Text
            string text = "Drag n drop files here or click to select ... \n\n.CSV, .TXT, .LOG, .XLSX";
            using (Font font = new Font("Segoe UI", 13F))
            using (Brush brush = new SolidBrush(Color.Gray))
            {
                SizeF textSize = g.MeasureString(text, font);

                float x = (dragDropPanel.Width / 2);
                float y = (dragDropPanel.Height / 2);
                g.DrawString(text, font, brush, x, y, new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center });
            }
        }

        private void DragDropPanel_DragDrop(object sender, DragEventArgs e)
        {
            dragDropPanel.BackColor = Color.LightGreen;
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            HandleDroppedFiles(files);
        }

        private void DragDropPanel_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
                dragDropPanel.BackColor = Color.DarkGray;
            }

        }

        private void MainForm_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            HandleDroppedFiles(files);
        }

        private void MainForm_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
        }

        private void btClearLog_Click(object sender, EventArgs e)
        {
            logTextBox.Clear();
        }

        private void AddLog(string message, Color? color = null)
        {
            if (logTextBox.InvokeRequired)
            {
                logTextBox.Invoke(new Action(() => AddLog(message, color)));
                return;
            }

            string timestamp = DateTime.Now.ToString("HH:mm:ss");
            logTextBox.SelectionStart = logTextBox.TextLength;
            logTextBox.SelectionLength = 0;
            logTextBox.SelectionColor = Color.Gray;
            logTextBox.AppendText($"[{timestamp}] ");
            logTextBox.SelectionColor = color ?? Color.Black;
            logTextBox.AppendText(message + "\n");
            logTextBox.SelectionColor = logTextBox.ForeColor;
            logTextBox.ScrollToCaret();
        }

        private void HandleDroppedFiles(string[] files)
        {
            currentFiles.Clear();
            foreach (string file in files)
            {
                string ext = Path.GetExtension(file).ToLower();
                if (ext == ".csv" || ext == ".txt" || ext == ".xlsx" || ext == ".log")
                {
                    currentFiles.Add(file);
                    AddLog($"Ready to convert: {file}");
                }
            }

            if (currentFiles.Count > 0)
            {
                convertButton.Enabled = true;

                UpdateStatusLabel($"{currentFiles.Count} file(s) selected");
            }
            else
            {
                MessageBox.Show("No supportet files found.\n\nsupportet formats: .csv, .txt, .log, .xlsx",
                    "DataWizard", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void UpdateStatusLabel(string text)
        {
            var statusStrip = this.Controls.OfType<StatusStrip>().FirstOrDefault();
            if (statusStrip != null)
            {
                var label = statusStrip.Items["statusLabel"] as ToolStripStatusLabel;
                if (label != null)
                {
                    label.Text = text;
                }
            }
        }

        private void dragDropPanel_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Excel Dateien|*.xlsx|CSV/Text Dateien|*.csv;*.txt;*.log|Alle Dateien|*.*";
                ofd.Multiselect = true;
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    HandleDroppedFiles(ofd.FileNames);
                }

            }
        }

        private void convertButton_Click(object sender, EventArgs e)
        {
            Convert();
        }

        private void Convert()
        {
            if (currentFiles.Count == 0) return;

            convertButton.Enabled = false;
            progressBar.Maximum = currentFiles.Count;
            progressBar.Value = 0;

            BackgroundWorker worker = new BackgroundWorker();
            worker.WorkerReportsProgress = true;
            worker.DoWork += Worker_DoWork;
            worker.ProgressChanged += Worker_ProgressChanged;
            worker.RunWorkerCompleted += Worker_RunWorkerCompleted;
            worker.RunWorkerAsync(currentFiles.ToList());
        }

        private void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //convertButton.Enabled = true;
            progressBar.Value = currentFiles.Count;
            UpdateStatusLabel("Job finished");
            currentFiles.Clear();
            // MessageBox.Show("Konvertierung abgeschlossen!", "DataWizard", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void Worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar.Value = e.ProgressPercentage;
            string message = e.UserState as string;
            if (message.StartsWith("✓"))
            {
                AddLog(message, Color.Green);
            }
            else
            {
                AddLog(message, Color.Red);
            }
        }

        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            List<string> files = e.Argument as List<string>;

            int processed = 0;
            foreach (string file in files)
            {
                try
                {
                    string extension = Path.GetExtension(file).ToLower();
                    string outputFile = GetOutputFileName(file, extension == ".xlsx" ? ".csv" : ".xlsx");

                    if (extension == ".csv" || extension == ".txt" || extension == ".log")
                    {
                        ConvertCsvToXlsx(file, outputFile);
                    }
                    else if (extension == ".xlsx")
                    {
                        ConvertXlsxToCsv(file, outputFile);
                    }

                    processed++;
                    worker.ReportProgress(processed, $"✓ {Path.GetFileName(file)} → {Path.GetFileName(outputFile)}");
                }
                catch (Exception ex)
                {
                    worker.ReportProgress(processed, $"✗ Fehler bei {Path.GetFileName(file)}: {ex.Message}");
                }
            }
        }

        private string GetOutputFileName(string inputPath, string newExtension)
        {
            string outDir = string.IsNullOrWhiteSpace(outputFolder)
                ? Path.GetDirectoryName(inputPath)
                : outputFolder;

            string baseFileName = Path.GetFileNameWithoutExtension(inputPath);
            string outputFile = Path.Combine(outDir, baseFileName + newExtension);

            if (!isOverwrite && File.Exists(outputFile))
            {
                int counter = 1;
                while (File.Exists(outputFile))
                {
                    outputFile = Path.Combine(outDir, $"{baseFileName}_{counter}{newExtension}");
                    counter++;
                }
            }

            return outputFile;
        }

        private void ConvertCsvToXlsx(string csvFile, string xlsxFile)
        {
            AddLog($"Konvertiere CSV → XLSX: {csvFile}");

            CSV csv = new CSV();
            csv.Load(csvFile, optionCsvSeparator, optionCsvEncoding, optionCsvHeader);

            AddLog($"Encoding: {csv.DetectedEncoding.EncodingName}");
            AddLog($"Separator: '{csv.Separator}'");
            AddLog($"Separator-Probability: {csv.SeparatorProbability:F2}%");
            AddLog($"Start-Line: {csv.StartLine}");
            AddLog($"End-Line: {csv.EndLine}");
            AddLog($"Lines: {csv.Lines}");
            AddLog($"Fields: {csv.FieldCount}");
            AddLog($"FieldCountConsistence: {csv.IsFieldCountEqual}");
            AddLog($"Header detected: {csv.DetectedHeaderLine}");

            AddLog($"Create: {xlsxFile}");
            csv.WriteXLSX(xlsxFile, isOverwrite);
        }

        private void ConvertXlsxToCsv(string xlsxFile, string csvFile)
        {
            AddLog($"Konvertiere XLSX → CSV: {xlsxFile} → {csvFile}");
            XLS.ToCsv(xlsxFile, csvFile, optionXlsSeparator,
                 optionXlsEncoding,
                 optionXlsQuoteAllText,               
                 0,
                 optionXlsAllSheets);


        }

        private void cbOverwrite_CheckedChanged(object sender, EventArgs e)
        {
            isOverwrite = cbOverwrite.Checked;
            saveSettings();
        }

        private void saveSettings()
        {
            // TODO
        }

        private void tbOutputfolder_TextChanged(object sender, EventArgs e)
        {
            outputFolder = tbOutputfolder.Text;
        }

        private Encoding getSelectedEncoding(string text)
        {
            switch (text)
            {
                case "UTF-8":
                    return Encoding.UTF8;
                case "UTF-8 mit BOM":
                    return new UTF8Encoding(true);
                case "Windows-1252":
                    return Encoding.GetEncoding(1252);
                case "ISO-8859-1":
                    return Encoding.GetEncoding("ISO-8859-1");
                default:
                    return null;
            }

        }

        private char getSelectedSeparator(string text)
        {
            switch (text)
            {
                case "Semicolon (;)":
                    return ';';
                case "Comma (,)":
                    return ',';
                case "Tab":
                    return '\t';
                case "Pipe (|)":
                    return '|';
                default:
                    return '\0';
            }

        }

        private void cbCsvEncoding_SelectedIndexChanged(object sender, EventArgs e)
        {
            optionCsvEncoding = null;
            if (cbCsvEncoding.SelectedItem.ToString() != "Auto")
            {
                optionCsvEncoding = getSelectedEncoding(cbCsvEncoding.SelectedItem.ToString());
            }
        }

        private void cbCsvSeparator_SelectedIndexChanged(object sender, EventArgs e)
        {
            optionCsvSeparator = '\0';
            if (cbCsvSeparator.SelectedItem.ToString() != "Auto")
            {
                optionCsvSeparator = getSelectedSeparator(cbCsvSeparator.SelectedItem.ToString());
            }
        }

        private void cbHeader_SelectedIndexChanged(object sender, EventArgs e)
        {
            optionCsvHeader = -1;
            if (cbHeader.SelectedItem.ToString() != "Auto")
            {
                optionCsvHeader = cbHeader.SelectedItem.ToString() == "Header" ? 1 : 0;
            }
        }

        private void cbQuoteAllText_CheckedChanged(object sender, EventArgs e)
        {
            optionXlsQuoteAllText = cbQuoteAllText.Checked;

        }

        private void cbAllSheets_CheckedChanged(object sender, EventArgs e)
        {
            optionXlsAllSheets = cbAllSheets.Checked;   
        }

        private void cbXlsSeparator_SelectedIndexChanged(object sender, EventArgs e)
        {
            optionXlsSeparator = getSelectedSeparator(cbXlsSeparator.SelectedItem.ToString());              
        }

        private void cbXlsEncoding_SelectedIndexChanged(object sender, EventArgs e)
        {
            optionXlsEncoding = getSelectedEncoding(cbXlsEncoding.SelectedItem.ToString());
        }
    }
}
