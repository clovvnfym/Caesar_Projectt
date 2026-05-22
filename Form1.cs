using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Text;
using System.Collections.Generic;
using Xceed.Words.NET;

namespace CaesarProject
{
    public partial class Form1 : Form
    {
        private TextBox? txtInput;
        private TextBox? txtOutput;
        private NumericUpDown? numShift;
        private Button? btnEncrypt;
        private Button? btnDecrypt;
        private Button? btnCopy;
        private Button? btnFileLoad;
        private Button? btnFileSave;
        private Button? btnClear;
        private TabControl? tabs;
        private TabPage? tab1;
        private TabPage? tab2;
        private TabPage? tab3;
        private ListBox? lstBatchFiles;
        private Button? btnAddBatch;
        private Button? btnClearBatch;
        private Button? btnProcessBatch;
        private NumericUpDown? numBatchShift;
        private ComboBox? cmbBatchMode;

        private string fileFilter = "Все поддерживаемые|*.txt;*.docx|Текстовые файлы (*.txt)|*.txt|Документы Word (*.docx)|*.docx";

        public Form1()
        {
            InitializeCustomComponents();
        }

        private void InitializeCustomComponents()
        {
            // Настройки окна
            this.Text = "Caesar Cipher Premium v1.0";
            this.Size = new Size(460, 730);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(25, 25, 25);
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;

            // Названия и параметры вкладок
            tabs = new TabControl { Dock = DockStyle.Fill, Appearance = TabAppearance.FlatButtons };
            tab1 = new TabPage { Text = "📝 ТЕКСТ", BackColor = Color.FromArgb(25, 25, 25) };
            tab2 = new TabPage { Text = "📁 ФАЙЛЫ", BackColor = Color.FromArgb(25, 25, 25), AllowDrop = true };
            tab3 = new TabPage { Text = "❓ СПРАВКА", BackColor = Color.FromArgb(25, 25, 25) };
            tabs.TabPages.AddRange(new TabPage[] { tab1, tab2, tab3 });
            this.Controls.Add(tabs);

            // Шрифт
            Font labelFont = new Font("Segoe UI", 10, FontStyle.Regular);

            // ВКЛАДКА 1
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
            btnCopy.Size = new Size(310, 45);

            btnFileSave = CreateStyledButton("💾 СОХРАНИТЬ В ФАЙЛ", new Point(25, 560), Color.FromArgb(45, 45, 45));
            btnFileSave.Size = new Size(310, 45);

            btnClear = CreateStyledButton("🗑", new Point(340, 505), Color.FromArgb(80, 40, 40));
            btnClear.Size = new Size(70, 100);
            btnClear.Font = new Font("Segoe UI Emoji", 12, FontStyle.Regular);

            tab1.Controls.AddRange(new Control[] { lblTitle, btnFileLoad, lblInput, txtInput, lblShift, numShift, btnEncrypt, btnDecrypt, lblResult, txtOutput, btnCopy, btnFileSave, btnClear });

            // ВКЛАДКА 2
            Label lblBatchTitle = new Label { Text = "ОБРАБОТКА ФАЙЛОВ", ForeColor = Color.FromArgb(0, 200, 150), Font = new Font("Segoe UI Semibold", 18, FontStyle.Bold), AutoSize = true, Location = new Point(25, 15) };
            lstBatchFiles = new ListBox { Location = new Point(25, 65), Size = new Size(395, 300), BackColor = Color.FromArgb(35, 35, 35), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle, Font = new Font("Segoe UI", 9), AllowDrop = true };

            Label lblBatchShift = new Label { Text = "Ключ сдвига:", ForeColor = Color.DarkGray, Font = labelFont, Location = new Point(25, 380), AutoSize = true };
            numBatchShift = new NumericUpDown { Location = new Point(140, 378), Width = 80, Minimum = -25, Maximum = 25, Value = 3, BackColor = Color.FromArgb(35, 35, 35), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle, TextAlign = HorizontalAlignment.Center };
            // Выпадающий список выбора режима
            cmbBatchMode = new ComboBox
            {
                Location = new Point(235, 378),
                Size = new Size(185, 30),
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = Color.FromArgb(35, 35, 35),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            };
            cmbBatchMode.Items.AddRange(new string[] { "🔐 ЗАШИФРОВАТЬ", "🔓 РАСШИФРОВАТЬ" });
            cmbBatchMode.SelectedIndex = 0; 

            btnAddBatch = CreateStyledButton("➕ ДОБАВИТЬ", new Point(25, 420), Color.FromArgb(100, 45, 140));
            btnAddBatch.Size = new Size(192, 45);

            btnClearBatch = CreateStyledButton("🗑 ОЧИСТИТЬ", new Point(228, 420), Color.FromArgb(80, 40, 40));
            btnClearBatch.Size = new Size(192, 45);

            btnProcessBatch = CreateStyledButton("🚀 НАЧАТЬ ОБРАБОТКУ", new Point(25, 480), Color.FromArgb(0, 110, 210));
            btnProcessBatch.Size = new Size(395, 60);

            Label lblInfo = new Label { Text = "Файлы (txt/docx) будут перезаписаны новым текстом.", ForeColor = Color.Gray, Location = new Point(25, 550), AutoSize = true, Font = new Font("Segoe UI", 8) };

           tab2.Controls.AddRange(new Control[] { lblBatchTitle, lstBatchFiles, lblBatchShift, numBatchShift, cmbBatchMode, btnAddBatch, btnClearBatch, btnProcessBatch, lblInfo });


            // ВКЛАДКА 3
            Label HelpTitle = new Label
            {
                Text = "СПРАВКА",
                ForeColor = Color.FromArgb(255, 180, 0),
                Font = new Font("Segoe UI Semibold", 20, FontStyle.Bold),
                AutoSize = true,
                Location = new Point(25, 15)
            };

            RichTextBox rtbHelp = new RichTextBox
            {
                Location = new Point(25, 65),
                Size = new Size(395, 530),
                BackColor = Color.FromArgb(35, 35, 35),
                ForeColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Georgia", 10),
                ReadOnly = true,
                ScrollBars = RichTextBoxScrollBars.Vertical,
                Text =
                    "📝 ВКЛАДКА «ТЕКСТ»\r\n" +
                    "--------------------------------------\r\n" +
                    "• Загрузить - загружает текст из файла TXT или DOCX.\r\n\r\n" +
                    "• Ключ сдвига - число от -25 до 25, на которое сдвигаются буквы.\r\n\r\n" +
                    "• ЗАШИФРОВАТЬ / РАСШИФРОВАТЬ - применяет шифр к тексту.\r\n\r\n" +
                    "• КОПИРОВАТЬ - копирует результат в буфер обмена.\r\n\r\n" +
                    "• СОХРАНИТЬ - сохраняет результат в файл TXT или DOCX.\r\n\r\n" +
                    "• 🗑 - очищает оба поля текста.\r\n\r\n\r\n" +
                    "📁 ВКЛАДКА «ФАЙЛЫ»\r\n" +
                    "--------------------------------------\r\n" +
                    "• Добавьте файлы TXT/DOCX (или перетащите их в список).\r\n\r\n" +
                    "• Укажите ключ сдвига и нажмите «НАЧАТЬ ОБРАБОТКУ».\r\n\r\n" +
                    "• Внимание: содержимое файлов будет перезаписано!\r\n\r\n\r\n" +
                    "💻 КОНСОЛЬНЫЙ РЕЖИМ\r\n" +
                    "--------------------------------------\r\n" +
                    "1. Открой PowerShell или cmd.\r\n\r\n" +
                    "2. Укажи путь к программе:\r\n" +
                    "   cd \"C:\\путь\\до\\папки\\с\\программой\"\r\n\r\n" +
                    "3. Зашифровать текст:\r\n" +
                    "   .\\CaesarProject.exe e 3 \"Hello\" -> Khoor\r\n\r\n" +
                    "4. Расшифровать текст:\r\n" +
                    "   .\\CaesarProject.exe d 3 \"Khoor\" -> Hello\r\n\r\n" +
                    "5. Зашифровать файл (файл будет перезаписан):\r\n" +
                    "   .\\CaesarProject.exe e 3 \"C:\\docs\\file.txt\"\r\n\r\n" +
                    "6. Расшифровать файл:\r\n" +
                    "   .\\CaesarProject.exe d 3 \"C:\\docs\\file.txt\"\r\n\r\n" +
                    "e - зашифровать, d - расшифровать.\r\n\r\n\r\n" +
                    "ℹ️ О ШИФРЕ ЦЕЗАРЯ\r\n" +
                    "--------------------------------------\r\n" +
                    "Каждая буква заменяется буквой на N позиций дальше в алфавите.\r\n" +
                    "Поддерживаются латинский и русский алфавиты.\r\n" +
                    "Цифры и знаки препинания не изменяются.\r\n\r\n" +
                    "Пример (сдвиг 3): А -> Г, B -> E, Hello -> Khoor"
            };

            tab3.Controls.AddRange(new Control[] { HelpTitle, rtbHelp });

            // СОБЫТИЯ

            tab2.DragEnter += (s, e) => { if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy; };
            tab2.DragDrop += (s, e) => HandleFileDrop(e);
            lstBatchFiles.DragEnter += (s, e) => { if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy; };
            lstBatchFiles.DragDrop += (s, e) => HandleFileDrop(e);

            btnEncrypt.Click += (s, e) =>
            {
                if (txtInput == null || string.IsNullOrWhiteSpace(txtInput.Text))
                {
                    MessageBox.Show("Введите текст для шифрования!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                if (numShift != null && numShift.Value == 0)
                {
                    MessageBox.Show("Сдвиг равен 0 — текст останется без изменений.", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                try
                {
                    if (txtOutput != null && numShift != null)
                        txtOutput.Text = CaesarCipher.Apply(txtInput.Text, (int)numShift.Value);
                }
                catch (Exception ex) { MessageBox.Show("Ошибка шифрования: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            };

            btnDecrypt.Click += (s, e) =>
            {
                if (txtInput == null || string.IsNullOrWhiteSpace(txtInput.Text))
                {
                    MessageBox.Show("Введите текст для расшифровки!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                if (numShift != null && numShift.Value == 0)
                {
                    MessageBox.Show("Сдвиг равен 0 — текст останется без изменений.", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                try
                {
                    if (txtOutput != null && numShift != null)
                        txtOutput.Text = CaesarCipher.Apply(txtInput.Text, -(int)numShift.Value);
                }
                catch (Exception ex) { MessageBox.Show("Ошибка расшифровки: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            };

            btnFileLoad.Click += (s, e) =>
            {
                try
                {
                    OpenFileDialog ofd = new OpenFileDialog { Filter = fileFilter };
                    if (ofd.ShowDialog() == DialogResult.OK && txtInput != null)
                        txtInput.Text = SmartRead(ofd.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка загрузки файла: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };

            btnFileSave.Click += (s, e) =>
            {
                if (txtOutput == null || string.IsNullOrEmpty(txtOutput.Text))
                {
                    MessageBox.Show("Нет текста для сохранения!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                try
                {
                    SaveFileDialog sfd = new SaveFileDialog { Filter = fileFilter };
                    if (sfd.ShowDialog() == DialogResult.OK)
                        SmartSave(sfd.FileName, txtOutput.Text);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка сохранения: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };

            btnCopy.Click += (s, e) =>
            {
                if (txtOutput == null || string.IsNullOrEmpty(txtOutput.Text))
                {
                    MessageBox.Show("Нет текста для копирования!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                try
                {
                    Clipboard.SetText(txtOutput.Text);
                    btnCopy.Text = "✔️ ТЕКСТ СКОПИРОВАН";
                    System.Windows.Forms.Timer t = new System.Windows.Forms.Timer { Interval = 2000 };
                    t.Tick += (ss, ee) => { btnCopy.Text = "КОПИРОВАТЬ В БУФЕР"; t.Stop(); };
                    t.Start();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка копирования: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };

            btnAddBatch.Click += (s, e) =>
            {
                try
                {
                    OpenFileDialog ofd = new OpenFileDialog { Filter = fileFilter, Multiselect = true };
                    if (ofd.ShowDialog() == DialogResult.OK && lstBatchFiles != null)
                        foreach (string f in ofd.FileNames)
                            if (!lstBatchFiles.Items.Contains(f)) lstBatchFiles.Items.Add(f);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка добавления файлов: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };

            btnClearBatch.Click += (s, e) => lstBatchFiles?.Items.Clear();

            btnProcessBatch.Click += (s, e) =>
            {
                if (lstBatchFiles == null || lstBatchFiles.Items.Count == 0 || numBatchShift == null)
                {
                    MessageBox.Show("Выберите файлы для обработки!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                var confirm = MessageBox.Show(
                    $"Файлы ({lstBatchFiles.Items.Count} шт.) будут перезаписаны. Продолжить?",
                    "Подтверждение",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);
                if (confirm != DialogResult.Yes) return;

                try
                {
                    int shift = (int)numBatchShift.Value;
                    if (cmbBatchMode != null && cmbBatchMode.SelectedIndex == 1) 
                    {
                        shift = -shift;
                    }
                    foreach (string path in lstBatchFiles.Items)
                    {
                        string content = SmartRead(path);
                        SmartSave(path, CaesarCipher.Apply(content, shift));
                    }

                    MessageBox.Show($"Готово! Все файлы обработаны с шагом {shift}.");
                    lstBatchFiles.Items.Clear();
                }
                catch (Exception ex) { MessageBox.Show("Ошибка обработки: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            };

            btnClear.Click += (s, e) =>
            {
                if (txtInput != null) txtInput.Text = "";
                if (txtOutput != null) txtOutput.Text = "";
            };
        }

        private string SmartRead(string path)
        {
            string ext = Path.GetExtension(path).ToLower();
            if (ext == ".docx")
            {
                using (DocX doc = DocX.Load(path))
                {
                    StringBuilder sb = new StringBuilder();
                    foreach (var para in doc.Paragraphs)
                        sb.AppendLine(para.Text);
                    return sb.ToString();
                }
            }
            return File.ReadAllText(path, Encoding.UTF8);
        }

        private void SmartSave(string path, string content)
        {
            string ext = Path.GetExtension(path).ToLower();
            if (ext == ".docx")
            {
                using (DocX doc = DocX.Create(path))
                {
                    string[] lines = content.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
                    foreach (string line in lines)
                    {
                        var para = doc.InsertParagraph(line);
                        if (line.StartsWith("\t"))
                            para.IndentationFirstLine = 1f;
                    }
                    doc.Save();
                }
            }
            else
            {
                File.WriteAllText(path, content, Encoding.UTF8);
            }
        }

        private void HandleFileDrop(DragEventArgs e)
        {
            if (lstBatchFiles == null) return;
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            foreach (string f in files)
            {
                string ext = Path.GetExtension(f).ToLower();
                if ((ext == ".txt" || ext == ".docx") && !lstBatchFiles.Items.Contains(f))
                    lstBatchFiles.Items.Add(f);
            }
        }

        private Button CreateStyledButton(string text, Point location, Color bgColor)
        {
            return new Button
            {
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
    }
}
