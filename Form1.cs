using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Text;
using System.Collections.Generic;
using Xceed.Words.NET; // НЕОБХОДИМО: NuGet пакет DocX

namespace CaesarProject
{
    public partial class Form1 : Form
    {
        // --- ЭЛЕМЕНТЫ УПРАВЛЕНИЯ ---
        private TextBox? txtInput;
        private TextBox? txtOutput;
        private NumericUpDown? numShift;
        private Button? btnEncrypt;
        private Button? btnDecrypt;
        private Button? btnCopy;
        private Button? btnFileLoad;
        private Button? btnFileSave;
        private TabControl? tabs;
        private TabPage? tab1;
        private TabPage? tab2;
        private ListBox? lstBatchFiles;
        private Button? btnAddBatch;
        private Button? btnClearBatch;
        private Button? btnProcessBatch;
        private NumericUpDown? numBatchShift;

        private string fileFilter = "Все поддерживаемые|*.txt;*.docx|Текстовые файлы (*.txt)|*.txt|Документы Word (*.docx)|*.docx";

        public Form1()
        {
            InitializeCustomComponents();
        }

        private void InitializeCustomComponents()
        {
            // --- НАСТРОЙКИ ОКНА ---
            this.Text = "Caesar Cipher Premium v2.0 (Word Support)";
            this.Size = new Size(460, 730);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(25, 25, 25);
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;

            // --- СИСТЕМА ВКЛАДОК ---
            tabs = new TabControl { Dock = DockStyle.Fill, Appearance = TabAppearance.FlatButtons };
            tab1 = new TabPage { Text = "📝 ТЕКСТ", BackColor = Color.FromArgb(25, 25, 25) };
            tab2 = new TabPage { Text = "📁 ФАЙЛЫ", BackColor = Color.FromArgb(25, 25, 25), AllowDrop = true };
            tabs.TabPages.AddRange(new TabPage[] { tab1, tab2 });
            this.Controls.Add(tabs);

            Font labelFont = new Font("Segoe UI", 10, FontStyle.Regular);

            // --- ВКЛАДКА 1: РАБОТА С ТЕКСТОМ ---
            Label lblTitle = new Label { Text = "ШИФР ЦЕЗАРЯ", ForeColor = Color.FromArgb(0, 150, 255), Font = new Font("Segoe UI Semibold", 20, FontStyle.Bold), AutoSize = true, Location = new Point(25, 15) };
            
            btnFileLoad = CreateStyledButton("📥 ЗАГРУЗИТЬ (TXT/DOCX)", new Point(25, 70), Color.FromArgb(100, 45, 140));
            btnFileLoad.Size = new Size(395, 40);

            Label lblInput = new Label { Text = "Исходный текст:", ForeColor = Color.DarkGray, Font = labelFont, Location = new Point(25, 120), AutoSize = true };
            txtInput = new TextBox { Multiline = true, Location = new Point(25, 145), Size = new Size(395, 100), BackColor = Color.FromArgb(35, 35, 35), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle, Font = new Font("Consolas", 11), ScrollBars = ScrollBars.Vertical };

            Label lblShift = new Label { Text = "Ключ сдвига:", ForeColor = Color.DarkGray, Font = labelFont, Location = new Point(25, 260), AutoSize = true };
            numShift = new NumericUpDown { Location = new Point(145, 258), Width = 80, Minimum = -25, Maximum = 25, Value = 3, BackColor = Color.FromArgb(35, 35, 35), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle, TextAlign = HorizontalAlignment.Center };

            btnEncrypt = CreateStyledButton("ЗАШИФРОВАТЬ", new Point(25, 300), Color.FromArgb(0, 110, 210));
            btnDecrypt = CreateStyledButton("РАСШИФРОВАТЬ", new Point(230, 300), Color.FromArgb(40, 130, 40));

            Label lblResult = new Label { Text = "Результат:", ForeColor = Color.DarkGray, Font = labelFont, Location = new Point(25, 360), AutoSize = true };
            txtOutput = new TextBox { Multiline = true, ReadOnly = true, Location = new Point(25, 385), Size = new Size(395, 100), BackColor = Color.FromArgb(35, 35, 35), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle, Font = new Font("Consolas", 11), ScrollBars = ScrollBars.Vertical };

            btnCopy = CreateStyledButton("КОПИРОВАТЬ В БУФЕР", new Point(25, 505), Color.FromArgb(60, 60, 60));
            btnCopy.Size = new Size(395, 45);

            btnFileSave = CreateStyledButton("💾 СОХРАНИТЬ В ФАЙЛ", new Point(25, 560), Color.FromArgb(45, 45, 45));
            btnFileSave.Size = new Size(395, 45);

            tab1.Controls.AddRange(new Control[] { lblTitle, btnFileLoad, lblInput, txtInput, lblShift, numShift, btnEncrypt, btnDecrypt, lblResult, txtOutput, btnCopy, btnFileSave });

            // --- ВКЛАДКА 2: ПАКЕТНАЯ ОБРАБОТКА (ОБНОВЛЕНО ПО СКРИНШОТУ) ---
            Label lblBatchTitle = new Label { Text = "ОБРАБОТКА ФАЙЛОВ", ForeColor = Color.FromArgb(0, 200, 150), Font = new Font("Segoe UI Semibold", 18, FontStyle.Bold), AutoSize = true, Location = new Point(25, 15) };
            
            lstBatchFiles = new ListBox { Location = new Point(25, 65), Size = new Size(395, 300), BackColor = Color.FromArgb(35, 35, 35), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle, Font = new Font("Segoe UI", 9), AllowDrop = true };

            Label lblBatchShift = new Label { Text = "Ключ сдвига:", ForeColor = Color.DarkGray, Font = labelFont, Location = new Point(25, 380), AutoSize = true };
            numBatchShift = new NumericUpDown { Location = new Point(140, 378), Width = 80, Minimum = -25, Maximum = 25, Value = 3, BackColor = Color.FromArgb(35, 35, 35), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle, TextAlign = HorizontalAlignment.Center };

            btnAddBatch = CreateStyledButton("➕ ДОБАВИТЬ", new Point(25, 420), Color.FromArgb(100, 45, 140));
            btnAddBatch.Size = new Size(192, 45);

            btnClearBatch = CreateStyledButton("🗑 ОЧИСТИТЬ", new Point(228, 420), Color.FromArgb(80, 40, 40));
            btnClearBatch.Size = new Size(192, 45);

            btnProcessBatch = CreateStyledButton("🚀 НАЧАТЬ ОБРАБОТКУ", new Point(25, 480), Color.FromArgb(0, 110, 210));
            btnProcessBatch.Size = new Size(395, 60);

            Label lblInfo = new Label { Text = "Файлы (txt/docx) будут перезаписаны новым текстом.", ForeColor = Color.Gray, Location = new Point(25, 550), AutoSize = true, Font = new Font("Segoe UI", 8) };

            tab2.Controls.AddRange(new Control[] { lblBatchTitle, lstBatchFiles, lblBatchShift, numBatchShift, btnAddBatch, btnClearBatch, btnProcessBatch, lblInfo });

            // --- ЛОГИКА СОБЫТИЙ ---
            tab2.DragEnter += (s, e) => { if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy; };
            tab2.DragDrop += (s, e) => HandleFileDrop(e);
            lstBatchFiles.DragEnter += (s, e) => { if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy; };
            lstBatchFiles.DragDrop += (s, e) => HandleFileDrop(e);

            btnEncrypt.Click += (s, e) => { if (txtInput != null && txtOutput != null && numShift != null) txtOutput.Text = ApplyCaesar(txtInput.Text, (int)numShift.Value); };
            btnDecrypt.Click += (s, e) => { if (txtInput != null && txtOutput != null && numShift != null) txtOutput.Text = ApplyCaesar(txtInput.Text, -(int)numShift.Value); };

            btnFileLoad.Click += (s, e) => {
                OpenFileDialog ofd = new OpenFileDialog { Filter = fileFilter };
                if (ofd.ShowDialog() == DialogResult.OK && txtInput != null) {
                    txtInput.Text = SmartRead(ofd.FileName);
                }
            };

            btnFileSave.Click += (s, e) => {
                if (txtOutput == null || string.IsNullOrEmpty(txtOutput.Text)) return;
                SaveFileDialog sfd = new SaveFileDialog { Filter = fileFilter };
                if (sfd.ShowDialog() == DialogResult.OK) {
                    SmartSave(sfd.FileName, txtOutput.Text);
                }
            };

            btnCopy.Click += (s, e) => {
                if (txtOutput != null && !string.IsNullOrEmpty(txtOutput.Text)) {
                    Clipboard.SetText(txtOutput.Text);
                    btnCopy.Text = "✔️ ТЕКСТ СКОПИРОВАН";
                    System.Windows.Forms.Timer t = new System.Windows.Forms.Timer { Interval = 2000 };
                    t.Tick += (ss, ee) => { btnCopy.Text = "КОПИРОВАТЬ В БУФЕР"; t.Stop(); };
                    t.Start();
                }
            };

            btnAddBatch.Click += (s, e) => {
                OpenFileDialog ofd = new OpenFileDialog { Filter = fileFilter, Multiselect = true };
                if (ofd.ShowDialog() == DialogResult.OK && lstBatchFiles != null) {
                    foreach (string f in ofd.FileNames)
                        if (!lstBatchFiles.Items.Contains(f)) lstBatchFiles.Items.Add(f);
                }
            };

            btnClearBatch.Click += (s, e) => lstBatchFiles?.Items.Clear();

            btnProcessBatch.Click += (s, e) => {
                if (lstBatchFiles == null || lstBatchFiles.Items.Count == 0 || numBatchShift == null) {
                    MessageBox.Show("Выберите файлы для обработки!");
                    return;
                }
                try {
                    int shift = (int)numBatchShift.Value;
                    foreach (string path in lstBatchFiles.Items) {
                        string content = SmartRead(path);
                        SmartSave(path, ApplyCaesar(content, shift));
                    }
                    MessageBox.Show($"Готово! Все файлы обработаны с шагом {shift}.");
                    lstBatchFiles.Items.Clear();
                }
                catch (Exception ex) { MessageBox.Show("Ошибка: " + ex.Message); }
            };
        }

        // --- МЕТОДЫ ДЛЯ WORD И ТЕКСТА ---
        private string SmartRead(string path) {
            string ext = Path.GetExtension(path).ToLower();
            if (ext == ".docx") {
                using (DocX doc = DocX.Load(path)) return doc.Text;
            }
            return File.ReadAllText(path, Encoding.Default);
        }

        private void SmartSave(string path, string content) {
            string ext = Path.GetExtension(path).ToLower();
            if (ext == ".docx") {
                using (DocX doc = DocX.Create(path)) {
                    doc.InsertParagraph(content);
                    doc.Save();
                }
            } else {
                File.WriteAllText(path, content, Encoding.UTF8);
            }
        }

        private void HandleFileDrop(DragEventArgs e) {
            if (lstBatchFiles == null) return;
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            foreach (string f in files) {
                string ext = Path.GetExtension(f).ToLower();
                if ((ext == ".txt" || ext == ".docx") && !lstBatchFiles.Items.Contains(f)) {
                    lstBatchFiles.Items.Add(f);
                }
            }
        }

        private Button CreateStyledButton(string text, Point location, Color bgColor) {
            return new Button {
                Text = text,
                Location = location,
                Size = new Size(190, 45),
                FlatStyle = FlatStyle.Flat,
                BackColor = bgColor,
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
        }

        private string ApplyCaesar(string input, int shift) {
            if (string.IsNullOrEmpty(input)) return string.Empty;
            StringBuilder res = new StringBuilder();
            foreach (char c in input) {
                if (c >= 'a' && c <= 'z') res.Append((char)('a' + (c - 'a' + (shift % 26) + 26) % 26));
                else if (c >= 'A' && c <= 'Z') res.Append((char)('A' + (c - 'A' + (shift % 26) + 26) % 26));
                else if (c >= 'а' && c <= 'я') res.Append((char)('а' + (c - 'а' + (shift % 32) + 32) % 32));
                else if (c >= 'А' && c <= 'Я') res.Append((char)('А' + (c - 'А' + (shift % 32) + 32) % 32));
                else res.Append(c);
            }
            return res.ToString();
        }
    }
}
