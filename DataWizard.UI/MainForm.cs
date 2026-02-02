using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DataWizard;
using libDataWizard;

namespace DataWizard.UI
{
    public partial class MainForm : Form
    {
        private Panel dragDropPanel;
        private TabControl mainTabControl;
        private TextBox outputFolderTextBox;
        private ComboBox encodingComboBox;
        private ComboBox separatorComboBox;
        private CheckBox autoDetectCheckBox;
        private CheckBox overwriteCheckBox;
        private RichTextBox logTextBox;
        private ProgressBar progressBar;
        private Button convertButton;
        private NotifyIcon trayIcon;
        private CheckBox watchModeCheckBox;
        private TextBox watchFolderTextBox;
        private FileSystemWatcher fileWatcher;
        private Timer watcherTimer;
        private bool isWatchMode = false;

        public MainForm()
        {
            InitializeComponent();
            InitializeCustomComponents();
            SetupTrayIcon();
            SetupFileWatcher();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            
            // Form Properties
            this.ClientSize = new Size(900, 650);
            this.Text = "DataWizard - CSV ↔ Excel Converter";
            this.MinimumSize = new Size(800, 600);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(240, 240, 240);
            this.Icon = SystemIcons.Application;

            // Allow Drag & Drop
            this.AllowDrop = true;
            this.DragEnter += MainForm_DragEnter;
            this.DragDrop += MainForm_DragDrop;

            this.ResumeLayout(false);
        }

        private void InitializeCustomComponents()
        {
            // Main TabControl
            mainTabControl = new TabControl
            {
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 9F)
            };
            this.Controls.Add(mainTabControl);

            // Tab 1: Konvertierung
            TabPage convertTab = new TabPage("Konvertierung");
            mainTabControl.TabPages.Add(convertTab);
            InitializeConvertTab(convertTab);

            // Tab 2: Einstellungen
            TabPage settingsTab = new TabPage("Einstellungen");
            mainTabControl.TabPages.Add(settingsTab);
            InitializeSettingsTab(settingsTab);

            // Tab 3: Überwachung
            TabPage watchTab = new TabPage("Ordner-Überwachung");
            mainTabControl.TabPages.Add(watchTab);
            InitializeWatchTab(watchTab);

            // Tab 4: Log
            TabPage logTab = new TabPage("Protokoll");
            mainTabControl.TabPages.Add(logTab);
            InitializeLogTab(logTab);

            // Status Bar
            StatusStrip statusStrip = new StatusStrip();
            ToolStripStatusLabel statusLabel = new ToolStripStatusLabel("Bereit");
            statusLabel.Name = "statusLabel";
            statusStrip.Items.Add(statusLabel);
            this.Controls.Add(statusStrip);
        }

        private void InitializeConvertTab(TabPage tab)
        {
            // Drag & Drop Panel
            dragDropPanel = new Panel
            {
                Location = new Point(20, 20),
                Size = new Size(840, 150),
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.White,
                AllowDrop = true
            };
            dragDropPanel.DragEnter += DragDropPanel_DragEnter;
            dragDropPanel.DragDrop += DragDropPanel_DragDrop;
            dragDropPanel.Paint += DragDropPanel_Paint;
            tab.Controls.Add(dragDropPanel);

            // File Selection Buttons
            Button selectCsvButton = new Button
            {
                Location = new Point(20, 190),
                Size = new Size(180, 40),
                Text = "CSV/TXT Datei wählen...",
                Font = new Font("Segoe UI", 9F),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White,
                Cursor = Cursors.Hand
            };
            selectCsvButton.FlatAppearance.BorderSize = 0;
            selectCsvButton.Click += SelectCsvButton_Click;
            tab.Controls.Add(selectCsvButton);

            Button selectXlsxButton = new Button
            {
                Location = new Point(220, 190),
                Size = new Size(180, 40),
                Text = "XLSX Datei wählen...",
                Font = new Font("Segoe UI", 9F),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(16, 124, 16),
                ForeColor = Color.White,
                Cursor = Cursors.Hand
            };
            selectXlsxButton.FlatAppearance.BorderSize = 0;
            selectXlsxButton.Click += SelectXlsxButton_Click;
            tab.Controls.Add(selectXlsxButton);

            // Output Folder
            Label outputLabel = new Label
            {
                Location = new Point(20, 250),
                Size = new Size(150, 20),
                Text = "Ausgabeordner:",
                Font = new Font("Segoe UI", 9F)
            };
            tab.Controls.Add(outputLabel);

            outputFolderTextBox = new TextBox
            {
                Location = new Point(20, 275),
                Size = new Size(650, 25),
                Font = new Font("Segoe UI", 9F),
                Text = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            };
            tab.Controls.Add(outputFolderTextBox);

            Button browseOutputButton = new Button
            {
                Location = new Point(680, 273),
                Size = new Size(100, 27),
                Text = "Durchsuchen...",
                Font = new Font("Segoe UI", 8F),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            browseOutputButton.Click += BrowseOutputButton_Click;
            tab.Controls.Add(browseOutputButton);

            // Progress Bar
            progressBar = new ProgressBar
            {
                Location = new Point(20, 320),
                Size = new Size(840, 25),
                Style = ProgressBarStyle.Continuous
            };
            tab.Controls.Add(progressBar);

            // Convert Button
            convertButton = new Button
            {
                Location = new Point(20, 360),
                Size = new Size(840, 50),
                Text = "Konvertieren",
                Font = new Font("Segoe UI", 11F, FontStyle.Bold),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White,
                Cursor = Cursors.Hand,
                Enabled = false
            };
            convertButton.FlatAppearance.BorderSize = 0;
            convertButton.Click += ConvertButton_Click;
            tab.Controls.Add(convertButton);

            // Info Label
            Label infoLabel = new Label
            {
                Location = new Point(20, 430),
                Size = new Size(840, 60),
                Text = "Ziehen Sie CSV/TXT oder XLSX Dateien hierher\noder wählen Sie Dateien über die Buttons aus.\n\nUnterstützte Formate: .csv, .txt, .xlsx",
                Font = new Font("Segoe UI", 9F),
                ForeColor = Color.Gray,
                TextAlign = ContentAlignment.TopLeft
            };
            tab.Controls.Add(infoLabel);
        }

        private void InitializeSettingsTab(TabPage tab)
        {
            GroupBox encodingGroup = new GroupBox
            {
                Location = new Point(20, 20),
                Size = new Size(400, 100),
                Text = "Zeichenkodierung",
                Font = new Font("Segoe UI", 9F)
            };
            tab.Controls.Add(encodingGroup);

            Label encodingLabel = new Label
            {
                Location = new Point(15, 30),
                Size = new Size(100, 20),
                Text = "Kodierung:",
                Font = new Font("Segoe UI", 9F)
            };
            encodingGroup.Controls.Add(encodingLabel);

            encodingComboBox = new ComboBox
            {
                Location = new Point(15, 55),
                Size = new Size(360, 25),
                Font = new Font("Segoe UI", 9F),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            encodingComboBox.Items.AddRange(new object[] 
            { 
                "Auto", 
                "UTF-8", 
                "UTF-8 mit BOM",
                "Windows-1252", 
                "ISO-8859-15",
                "ASCII"
            });
            encodingComboBox.SelectedIndex = 0;
            encodingGroup.Controls.Add(encodingComboBox);

            GroupBox separatorGroup = new GroupBox
            {
                Location = new Point(440, 20),
                Size = new Size(400, 100),
                Text = "CSV Trennzeichen",
                Font = new Font("Segoe UI", 9F)
            };
            tab.Controls.Add(separatorGroup);

            autoDetectCheckBox = new CheckBox
            {
                Location = new Point(15, 30),
                Size = new Size(350, 20),
                Text = "Automatisch erkennen",
                Font = new Font("Segoe UI", 9F),
                Checked = true
            };
            autoDetectCheckBox.CheckedChanged += AutoDetectCheckBox_CheckedChanged;
            separatorGroup.Controls.Add(autoDetectCheckBox);

            separatorComboBox = new ComboBox
            {
                Location = new Point(15, 55),
                Size = new Size(360, 25),
                Font = new Font("Segoe UI", 9F),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Enabled = false
            };
            separatorComboBox.Items.AddRange(new object[] 
            { 
                "Semikolon (;)", 
                "Komma (,)", 
                "Tabulator",
                "Pipe (|)"
            });
            separatorComboBox.SelectedIndex = 0;
            separatorGroup.Controls.Add(separatorComboBox);

            GroupBox fileHandlingGroup = new GroupBox
            {
                Location = new Point(20, 140),
                Size = new Size(820, 100),
                Text = "Dateibehandlung",
                Font = new Font("Segoe UI", 9F)
            };
            tab.Controls.Add(fileHandlingGroup);

            overwriteCheckBox = new CheckBox
            {
                Location = new Point(15, 30),
                Size = new Size(350, 20),
                Text = "Existierende Dateien überschreiben",
                Font = new Font("Segoe UI", 9F),
                Checked = false
            };
            fileHandlingGroup.Controls.Add(overwriteCheckBox);

            Label fileHandlingInfo = new Label
            {
                Location = new Point(15, 55),
                Size = new Size(780, 35),
                Text = "Wenn deaktiviert, wird bei vorhandenen Dateien automatisch ein Zähler angehängt.\nBeispiel: datei.xlsx → datei_1.xlsx, datei_2.xlsx, etc.",
                Font = new Font("Segoe UI", 8.5F),
                ForeColor = Color.Gray
            };
            fileHandlingGroup.Controls.Add(fileHandlingInfo);

            // Reset Button
            Button resetButton = new Button
            {
                Location = new Point(20, 260),
                Size = new Size(200, 35),
                Text = "Standardeinstellungen",
                Font = new Font("Segoe UI", 9F),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            resetButton.Click += ResetButton_Click;
            tab.Controls.Add(resetButton);
        }

        private void InitializeWatchTab(TabPage tab)
        {
            Label infoLabel = new Label
            {
                Location = new Point(20, 20),
                Size = new Size(840, 40),
                Text = "Überwachen Sie einen Ordner automatisch und konvertieren Sie neue Dateien sofort.\nDie Anwendung minimiert sich in den System-Tray.",
                Font = new Font("Segoe UI", 9F),
                ForeColor = Color.Gray
            };
            tab.Controls.Add(infoLabel);

            watchModeCheckBox = new CheckBox
            {
                Location = new Point(20, 75),
                Size = new Size(300, 25),
                Text = "Ordner-Überwachung aktivieren",
                Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                Checked = false
            };
            watchModeCheckBox.CheckedChanged += WatchModeCheckBox_CheckedChanged;
            tab.Controls.Add(watchModeCheckBox);

            GroupBox watchFolderGroup = new GroupBox
            {
                Location = new Point(20, 115),
                Size = new Size(840, 100),
                Text = "Überwachter Ordner",
                Font = new Font("Segoe UI", 9F)
            };
            tab.Controls.Add(watchFolderGroup);

            watchFolderTextBox = new TextBox
            {
                Location = new Point(15, 35),
                Size = new Size(650, 25),
                Font = new Font("Segoe UI", 9F),
                Enabled = false
            };
            watchFolderGroup.Controls.Add(watchFolderTextBox);

            Button browseWatchButton = new Button
            {
                Location = new Point(675, 33),
                Size = new Size(140, 27),
                Text = "Durchsuchen...",
                Font = new Font("Segoe UI", 8F),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                Enabled = false
            };
            browseWatchButton.Click += BrowseWatchButton_Click;
            watchFolderGroup.Controls.Add(browseWatchButton);

            Label watchInfo = new Label
            {
                Location = new Point(15, 68),
                Size = new Size(800, 20),
                Text = "Neue CSV/TXT und XLSX Dateien werden automatisch konvertiert.",
                Font = new Font("Segoe UI", 8.5F),
                ForeColor = Color.Gray
            };
            watchFolderGroup.Controls.Add(watchInfo);

            // Watch Status
            Label watchStatusLabel = new Label
            {
                Location = new Point(20, 230),
                Size = new Size(100, 20),
                Text = "Status:",
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };
            tab.Controls.Add(watchStatusLabel);

            Label watchStatusValue = new Label
            {
                Location = new Point(120, 230),
                Size = new Size(400, 20),
                Text = "Inaktiv",
                Font = new Font("Segoe UI", 9F),
                ForeColor = Color.Gray,
                Name = "watchStatusValue"
            };
            tab.Controls.Add(watchStatusValue);

            // Minimize to Tray Button
            Button minimizeButton = new Button
            {
                Location = new Point(20, 270),
                Size = new Size(300, 40),
                Text = "In System-Tray minimieren",
                Font = new Font("Segoe UI", 9F),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White,
                Cursor = Cursors.Hand,
                Enabled = false,
                Name = "minimizeButton"
            };
            minimizeButton.FlatAppearance.BorderSize = 0;
            minimizeButton.Click += MinimizeButton_Click;
            tab.Controls.Add(minimizeButton);
        }

        private void InitializeLogTab(TabPage tab)
        {
            logTextBox = new RichTextBox
            {
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 9F),
                BackColor = Color.White,
                ReadOnly = true,
                BorderStyle = BorderStyle.None
            };
            tab.Controls.Add(logTextBox);

            Panel buttonPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 50,
                Padding = new Padding(10)
            };
            tab.Controls.Add(buttonPanel);

            Button clearLogButton = new Button
            {
                Location = new Point(10, 10),
                Size = new Size(120, 30),
                Text = "Log löschen",
                Font = new Font("Segoe UI", 9F),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            clearLogButton.Click += (s, e) => logTextBox.Clear();
            buttonPanel.Controls.Add(clearLogButton);

            AddLog("DataWizard gestartet.");
        }

        private void SetupTrayIcon()
        {
            trayIcon = new NotifyIcon
            {
                Icon = SystemIcons.Application,
                Text = "DataWizard - Ordner-Überwachung",
                Visible = false
            };

            ContextMenuStrip trayMenu = new ContextMenuStrip();
            trayMenu.Items.Add("Anzeigen", null, (s, e) => ShowFromTray());
            trayMenu.Items.Add("Beenden", null, (s, e) => Application.Exit());
            trayIcon.ContextMenuStrip = trayMenu;
            trayIcon.DoubleClick += (s, e) => ShowFromTray();
        }

        private void SetupFileWatcher()
        {
            fileWatcher = new FileSystemWatcher();
            fileWatcher.NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite;
            fileWatcher.Created += FileWatcher_Created;
            fileWatcher.EnableRaisingEvents = false;

            // Timer für verzögerte Verarbeitung (Dateien komplett geschrieben)
            watcherTimer = new Timer();
            watcherTimer.Interval = 2000; // 2 Sekunden Verzögerung
            watcherTimer.Tick += WatcherTimer_Tick;
        }

        private Queue<string> pendingFiles = new Queue<string>();

        private void FileWatcher_Created(object sender, FileSystemEventArgs e)
        {
            string extension = Path.GetExtension(e.FullPath).ToLower();
            if (extension == ".csv" || extension == ".txt" || extension == ".xlsx")
            {
                lock (pendingFiles)
                {
                    if (!pendingFiles.Contains(e.FullPath))
                    {
                        pendingFiles.Enqueue(e.FullPath);
                        AddLog($"Neue Datei erkannt: {Path.GetFileName(e.FullPath)}");
                    }
                }
                
                if (!watcherTimer.Enabled)
                {
                    watcherTimer.Start();
                }
            }
        }

        private void WatcherTimer_Tick(object sender, EventArgs e)
        {
            watcherTimer.Stop();
            
            List<string> filesToProcess = new List<string>();
            lock (pendingFiles)
            {
                while (pendingFiles.Count > 0)
                {
                    filesToProcess.Add(pendingFiles.Dequeue());
                }
            }

            foreach (string file in filesToProcess)
            {
                if (File.Exists(file))
                {
                    ProcessFileInWatchMode(file);
                }
            }
        }

        private void ProcessFileInWatchMode(string filePath)
        {
            try
            {
                string extension = Path.GetExtension(filePath).ToLower();
                string outputFile = GetOutputFileName(filePath, extension == ".xlsx" ? ".csv" : ".xlsx");

                AddLog($"Konvertiere: {filePath}...");

                if (extension == ".csv" || extension == ".txt")
                {
                    ConvertCsvToXlsx(filePath, outputFile);
                }
                else if (extension == ".xlsx")
                {
                    ConvertXlsxToCsv(filePath, outputFile);
                }

                AddLog($"✓ Erfolgreich: {outputFile}", Color.Green);
            }
            catch (Exception ex)
            {
                AddLog($"✗ Fehler bei {filePath}: {ex.Message}", Color.Red);
            }
        }

        private string GetOutputFileName(string inputPath, string newExtension)
        {
            string outputFolder = string.IsNullOrWhiteSpace(outputFolderTextBox.Text) 
                ? Path.GetDirectoryName(inputPath) 
                : outputFolderTextBox.Text;

            string baseFileName = Path.GetFileNameWithoutExtension(inputPath);
            string outputFile = Path.Combine(outputFolder, baseFileName + newExtension);

            if (!overwriteCheckBox.Checked && File.Exists(outputFile))
            {
                int counter = 1;
                while (File.Exists(outputFile))
                {
                    outputFile = Path.Combine(outputFolder, $"{baseFileName}_{counter}{newExtension}");
                    counter++;
                }
            }

            return outputFile;
        }

        private void DragDropPanel_Paint(object sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.Clear(Color.White);

            // Gestrichelter Rahmen
            using (Pen dashedPen = new Pen(Color.FromArgb(0, 120, 215), 2))
            {
                dashedPen.DashStyle = System.Drawing.Drawing2D.DashStyle.Dash;
                g.DrawRectangle(dashedPen, 10, 10, dragDropPanel.Width - 20, dragDropPanel.Height - 20);
            }

            // Icon und Text
            string text = "Dateien hier ablegen\n\nCSV, TXT oder XLSX";
            using (Font font = new Font("Segoe UI", 12F))
            using (Brush brush = new SolidBrush(Color.Gray))
            {
                SizeF textSize = g.MeasureString(text, font);
                float x = (dragDropPanel.Width - textSize.Width) / 2;
                float y = (dragDropPanel.Height - textSize.Height) / 2;
                g.DrawString(text, font, brush, x, y, new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center });
            }
        }

        private void MainForm_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
        }

        private void MainForm_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            HandleDroppedFiles(files);
        }

        private void DragDropPanel_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
                dragDropPanel.BackColor = Color.FromArgb(230, 240, 255);
            }
        }

        private void DragDropPanel_DragDrop(object sender, DragEventArgs e)
        {
            dragDropPanel.BackColor = Color.White;
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            HandleDroppedFiles(files);
        }

        private List<string> currentFiles = new List<string>();

        private void HandleDroppedFiles(string[] files)
        {
            currentFiles.Clear();
            foreach (string file in files)
            {
                string ext = Path.GetExtension(file).ToLower();
                if (ext == ".csv" || ext == ".txt" || ext == ".xlsx")
                {
                    currentFiles.Add(file);
                }
            }

            if (currentFiles.Count > 0)
            {
                convertButton.Enabled = true;
                AddLog($"{currentFiles.Count} Datei(en) bereit zur Konvertierung.");
                UpdateStatusLabel($"{currentFiles.Count} Datei(en) ausgewählt");
            }
            else
            {
                MessageBox.Show("Keine unterstützten Dateien gefunden.\n\nUnterstützte Formate: .csv, .txt, .xlsx", 
                    "DataWizard", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void SelectCsvButton_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "CSV/Text Dateien|*.csv;*.txt|Alle Dateien|*.*";
                ofd.Multiselect = true;
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    HandleDroppedFiles(ofd.FileNames);
                }
            }
        }

        private void SelectXlsxButton_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Excel Dateien|*.xlsx|Alle Dateien|*.*";
                ofd.Multiselect = true;
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    HandleDroppedFiles(ofd.FileNames);
                }
            }
        }

        private void BrowseOutputButton_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog())
            {
                fbd.SelectedPath = outputFolderTextBox.Text;
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    outputFolderTextBox.Text = fbd.SelectedPath;
                }
            }
        }

        private void BrowseWatchButton_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog())
            {
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    watchFolderTextBox.Text = fbd.SelectedPath;
                }
            }
        }

        private void ConvertButton_Click(object sender, EventArgs e)
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

                    if (extension == ".csv" || extension == ".txt")
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

        private void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            convertButton.Enabled = true;
            progressBar.Value = 0;
            UpdateStatusLabel("Konvertierung abgeschlossen");
            MessageBox.Show("Konvertierung abgeschlossen!", "DataWizard", 
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ConvertCsvToXlsx(string csvFile, string xlsxFile)
        {
            // TODO: Implementierung mit libDataWizard
            AddLog($"Konvertiere CSV → XLSX: {csvFile}");
            
            CSV csv = new CSV();
            csv.Load(csvFile);
            
            // Hier kommt die XLS.WriteFromCSV Methode
            AddLog($"Erstelle: {xlsxFile}");
            csv.WriteXLSX(xlsxFile, true);
        }

        private void ConvertXlsxToCsv(string xlsxFile, string csvFile)
        {
            // TODO: Implementierung mit libDataWizard
                  
            AddLog($"Konvertiere XLSX → CSV: {xlsxFile} → {csvFile}");
            XLS.ToCsv(xlsxFile,csvFile);
        }

        private void AutoDetectCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            separatorComboBox.Enabled = !autoDetectCheckBox.Checked;
        }

        private void WatchModeCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            bool enabled = watchModeCheckBox.Checked;
            watchFolderTextBox.Enabled = enabled;
            
            var browseButton = mainTabControl.TabPages[2].Controls.Find("browseWatchButton", true).FirstOrDefault();
            if (browseButton != null) browseButton.Enabled = enabled;
            
            var minimizeButton = mainTabControl.TabPages[2].Controls.Find("minimizeButton", true).FirstOrDefault();
            if (minimizeButton != null) minimizeButton.Enabled = enabled;

            if (enabled && !string.IsNullOrWhiteSpace(watchFolderTextBox.Text))
            {
                StartWatchMode();
            }
            else
            {
                StopWatchMode();
            }
        }

        private void StartWatchMode()
        {
            try
            {
                fileWatcher.Path = watchFolderTextBox.Text;
                fileWatcher.Filter = "*.*";
                fileWatcher.EnableRaisingEvents = true;
                isWatchMode = true;

                var statusLabel = mainTabControl.TabPages[2].Controls.Find("watchStatusValue", true).FirstOrDefault() as Label;
                if (statusLabel != null)
                {
                    statusLabel.Text = $"Aktiv - Überwache: {watchFolderTextBox.Text}";
                    statusLabel.ForeColor = Color.Green;
                }

                AddLog($"Ordner-Überwachung gestartet: {watchFolderTextBox.Text}", Color.Blue);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fehler beim Starten der Überwachung:\n{ex.Message}", 
                    "DataWizard", MessageBoxButtons.OK, MessageBoxIcon.Error);
                watchModeCheckBox.Checked = false;
            }
        }

        private void StopWatchMode()
        {
            fileWatcher.EnableRaisingEvents = false;
            isWatchMode = false;

            var statusLabel = mainTabControl.TabPages[2].Controls.Find("watchStatusValue", true).FirstOrDefault() as Label;
            if (statusLabel != null)
            {
                statusLabel.Text = "Inaktiv";
                statusLabel.ForeColor = Color.Gray;
            }

            AddLog("Ordner-Überwachung beendet.", Color.Blue);
        }

        private void MinimizeButton_Click(object sender, EventArgs e)
        {
            this.Hide();
            trayIcon.Visible = true;
            trayIcon.ShowBalloonTip(2000, "DataWizard", 
                "Läuft im Hintergrund und überwacht den Ordner.", ToolTipIcon.Info);
        }

        private void ShowFromTray()
        {
            this.Show();
            this.WindowState = FormWindowState.Normal;
            this.Activate();
            trayIcon.Visible = false;
        }

        private void ResetButton_Click(object sender, EventArgs e)
        {
            encodingComboBox.SelectedIndex = 0;
            separatorComboBox.SelectedIndex = 0;
            autoDetectCheckBox.Checked = true;
            overwriteCheckBox.Checked = false;
            outputFolderTextBox.Text = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            AddLog("Einstellungen zurückgesetzt.");
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

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (isWatchMode && e.CloseReason == CloseReason.UserClosing)
            {
                DialogResult result = MessageBox.Show(
                    "Die Ordner-Überwachung ist aktiv.\n\nMöchten Sie die Anwendung wirklich beenden?",
                    "DataWizard", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                
                if (result == DialogResult.No)
                {
                    e.Cancel = true;
                    return;
                }
            }

            if (fileWatcher != null)
            {
                fileWatcher.EnableRaisingEvents = false;
                fileWatcher.Dispose();
            }

            if (trayIcon != null)
            {
                trayIcon.Visible = false;
                trayIcon.Dispose();
            }

            base.OnFormClosing(e);
        }
    }
}
