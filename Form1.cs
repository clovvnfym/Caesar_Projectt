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
        private TabPage? tab4;
        private Panel? eduPanel;
        private Label? lblEduResult;
        private TextBox? txtEduInput;
        private NumericUpDown? numEduShift;
        private ListBox? lstEduSteps;

        private string fileFilter = "Все поддерживаемые|*.txt;*.docx|Текстовые файлы (*.txt)|*.txt|Документы Word (*.docx)|*.docx";

        public Form1()
        {
            InitializeCustomComponents();
        }

        // ─────────────────────────────────────────────────────────────
        //  Проверка: является ли файл настоящим ZIP/DOCX
        // ─────────────────────────────────────────────────────────────
        private bool IsValidDocx(string path)
        {
            try
            {
                // DOCX — это ZIP. Первые 4 байта: 50 4B 03 04
                byte[] header = new byte[4];
                using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read))
                {
                    if (fs.Length < 4) return false;
                    fs.Read(header, 0, 4);
                }
                return header[0] == 0x50 && header[1] == 0x4B
                    && header[2] == 0x03 && header[3] == 0x04;
            }
            catch
            {
                return false;
            }
        }

        // ─────────────────────────────────────────────────────────────
        //  Чтение файла (TXT или DOCX)
        // ─────────────────────────────────────────────────────────────
        private string SmartRead(string path)
        {
            string ext = Path.GetExtension(path).ToLower();

            if (ext == ".docx")
            {
                // Сначала проверяем подпись ZIP
                if (!IsValidDocx(path))
                {
                    // Файл не является настоящим DOCX — пробуем прочитать как текст
                    try
                    {
                        return File.ReadAllText(path, Encoding.UTF8);
                    }
                    catch
                    {
                        throw new InvalidDataException(
                            "Файл имеет расширение .docx, но не является корректным документом Word.\n" +
                            "Возможно, файл повреждён или это переименованный TXT-файл.");
                    }
                }

                // Файл корректный — читаем через Xceed
                try
                {
                    using (DocX doc = DocX.Load(path))
                    {
                        StringBuilder sb = new StringBuilder();
                        foreach (var para in doc.Paragraphs)
                            sb.AppendLine(para.Text);

                        // Убираем нулевые байты и другие недопустимые XML-символы
                        return RemoveInvalidXmlChars(sb.ToString());
                    }
                }
                catch (Exception ex)
                {
                    throw new InvalidDataException(
                        "Не удалось открыть DOCX-файл.\n" +
                        "Детали: " + ex.Message);
                }
            }

            // TXT и любые другие расширения
            return File.ReadAllText(path, Encoding.UTF8);
        }
        private string RemoveInvalidXmlChars(string text)
        {
            StringBuilder result = new StringBuilder();
            foreach (char c in text)
            {
                if (c == 0x09 || c == 0x0A || c == 0x0D ||
                    (c >= 0x20 && c <= 0xD7FF) ||
                    (c >= 0xE000 && c <= 0xFFFD))
                {
                    result.Append(c);
                }
            }
            return result.ToString();
        }
        //  Сохранение файла (TXT или DOCX)
        private void SmartSave(string path, string content)
        {
            content = RemoveInvalidXmlChars(content);
            string ext = Path.GetExtension(path).ToLower();

            if (ext == ".docx")
            {
                using (DocX doc = DocX.Create(path))
                {
                    string[] lines = content.Split(
                        new[] { "\r\n", "\n" },
                        StringSplitOptions.None);

                    for (int i = 0; i < lines.Length; i++)
                    {
                        if (i == lines.Length - 1 && string.IsNullOrEmpty(lines[i]))
                            break;

                        doc.InsertParagraph(lines[i]);
                    }

                    doc.Save();
                }
            }
            else
            {
                File.WriteAllText(path, content, Encoding.UTF8);
            }
        }
        //  Drag-and-drop для вкладки «Файлы»
        private void HandleFileDrop(DragEventArgs e)
        {
            if (lstBatchFiles == null) return;
            string[] files = (string[])e.Data!.GetData(DataFormats.FileDrop);
            foreach (string f in files)
            {
                string ext = Path.GetExtension(f).ToLower();
                if ((ext == ".txt" || ext == ".docx") && !lstBatchFiles.Items.Contains(f))
                    lstBatchFiles.Items.Add(f);
            }
        }

        //  Вспомогательный метод создания кнопки
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
        //  Построение интерфейса
        private void InitializeCustomComponents()
        {
            this.Text = "Caesar Cipher Premium v1.0";
            this.Size = new Size(460, 730);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(25, 25, 25);
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;

            tabs = new TabControl { Dock = DockStyle.Fill, Appearance = TabAppearance.FlatButtons };
            tab1 = new TabPage { Text = "📝 ТЕКСТ",   BackColor = Color.FromArgb(25, 25, 25) };
            tab2 = new TabPage { Text = "📁 ФАЙЛЫ",   BackColor = Color.FromArgb(25, 25, 25), AllowDrop = true };
            tab3 = new TabPage { Text = "❓ СПРАВКА", BackColor = Color.FromArgb(25, 25, 25) };
            tabs.TabPages.AddRange(new TabPage[] { tab1, tab2, tab3 });
            this.Controls.Add(tabs);

            Font labelFont = new Font("Segoe UI", 10, FontStyle.Regular);

            //  ВКЛАДКА 1 
            Label lblTitle = new Label
            {
                Text = "ШИФР ЦЕЗАРЯ",
                ForeColor = Color.FromArgb(0, 150, 255),
                Font = new Font("Segoe UI Semibold", 20, FontStyle.Bold),
                AutoSize = true,
                Location = new Point(25, 15)
            };

            btnFileLoad = CreateStyledButton("📥 ЗАГРУЗИТЬ (TXT/DOCX)", new Point(25, 70), Color.FromArgb(100, 45, 140));
            btnFileLoad.Size = new Size(395, 40);

            Label lblInput = new Label
            {
                Text = "Исходный текст:",
                ForeColor = Color.DarkGray,
                Font = labelFont,
                Location = new Point(25, 120),
                AutoSize = true
            };

            txtInput = new TextBox
            {
                Multiline = true,
                Location = new Point(25, 145),
                Size = new Size(395, 100),
                BackColor = Color.FromArgb(35, 35, 35),
                ForeColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Consolas", 11),
                ScrollBars = ScrollBars.Vertical
            };

            Label lblShift = new Label
            {
                Text = "Ключ сдвига:",
                ForeColor = Color.DarkGray,
                Font = labelFont,
                Location = new Point(25, 260),
                AutoSize = true
            };

            numShift = new NumericUpDown
            {
                Location = new Point(145, 258),
                Width = 80,
                Minimum = -25,
                Maximum = 25,
                Value = 3,
                BackColor = Color.FromArgb(35, 35, 35),
                ForeColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                TextAlign = HorizontalAlignment.Center
            };

            btnEncrypt = CreateStyledButton("ЗАШИФРОВАТЬ", new Point(25, 300), Color.FromArgb(60, 60, 60));
            btnDecrypt = CreateStyledButton("РАСШИФРОВАТЬ", new Point(230, 300), Color.FromArgb(60, 60, 60));

            Label lblResult = new Label
            {
                Text = "Результат:",
                ForeColor = Color.DarkGray,
                Font = labelFont,
                Location = new Point(25, 360),
                AutoSize = true
            };

            txtOutput = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                Location = new Point(25, 385),
                Size = new Size(395, 100),
                BackColor = Color.FromArgb(35, 35, 35),
                ForeColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Consolas", 11),
                ScrollBars = ScrollBars.Vertical
            };

            btnCopy = CreateStyledButton("КОПИРОВАТЬ В БУФЕР", new Point(25, 505), Color.FromArgb(60, 60, 60));
            btnCopy.Size = new Size(310, 45);

            btnFileSave = CreateStyledButton("💾 СОХРАНИТЬ В ФАЙЛ", new Point(25, 560), Color.FromArgb(45, 45, 45));
            btnFileSave.Size = new Size(310, 45);

            btnClear = CreateStyledButton("🗑", new Point(340, 505), Color.FromArgb(80, 40, 40));
            btnClear.Size = new Size(70, 100);
            btnClear.Font = new Font("Segoe UI Emoji", 12, FontStyle.Regular);

            tab1.Controls.AddRange(new Control[]
            {
                lblTitle, btnFileLoad, lblInput, txtInput,
                lblShift, numShift, btnEncrypt, btnDecrypt,
                lblResult, txtOutput, btnCopy, btnFileSave, btnClear
            });

            //  ВКЛАДКА 2 
            Label lblBatchTitle = new Label
            {
                Text = "ОБРАБОТКА ФАЙЛОВ",
                ForeColor = Color.FromArgb(0, 200, 150),
                Font = new Font("Segoe UI Semibold", 18, FontStyle.Bold),
                AutoSize = true,
                Location = new Point(25, 15)
            };

            lstBatchFiles = new ListBox
            {
                Location = new Point(25, 65),
                Size = new Size(395, 300),
                BackColor = Color.FromArgb(35, 35, 35),
                ForeColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 9),
                AllowDrop = true
            };

            Label lblBatchShift = new Label
            {
                Text = "Ключ сдвига:",
                ForeColor = Color.DarkGray,
                Font = labelFont,
                Location = new Point(25, 380),
                AutoSize = true
            };

            numBatchShift = new NumericUpDown
            {
                Location = new Point(140, 378),
                Width = 80,
                Minimum = -25,
                Maximum = 25,
                Value = 3,
                BackColor = Color.FromArgb(35, 35, 35),
                ForeColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                TextAlign = HorizontalAlignment.Center
            };

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
            cmbBatchMode.Items.AddRange(new string[] { "🔒 ЗАШИФРОВАТЬ", "🔓 РАСШИФРОВАТЬ" });
            cmbBatchMode.SelectedIndex = 0;

            btnAddBatch = CreateStyledButton("➕ ДОБАВИТЬ", new Point(25, 420), Color.FromArgb(70, 70, 70));
            btnAddBatch.Size = new Size(192, 45);

            btnClearBatch = CreateStyledButton("🗑 ОЧИСТИТЬ", new Point(228, 420), Color.FromArgb(70, 70, 70));
            btnClearBatch.Size = new Size(192, 45);

            btnProcessBatch = CreateStyledButton("🚀 НАЧАТЬ ОБРАБОТКУ", new Point(25, 480), Color.FromArgb(90, 90, 90));
            btnProcessBatch.Size = new Size(395, 60);

            Label lblInfo = new Label
            {
                Text = "Файлы (txt/docx) будут перезаписаны новым текстом.",
                ForeColor = Color.Gray,
                Location = new Point(25, 550),
                AutoSize = true,
                Font = new Font("Segoe UI", 8)
            };

            tab2.Controls.AddRange(new Control[]
            {
                lblBatchTitle, lstBatchFiles, lblBatchShift, numBatchShift,
                cmbBatchMode, btnAddBatch, btnClearBatch, btnProcessBatch, lblInfo
            });

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
                    "e - зашифровать, d - расшифровать.\r\n\r\n"+
                    "ℹ️ О ШИФРЕ ЦЕЗАРЯ\r\n" +
                    "--------------------------------------\r\n" +
                    "Каждая буква заменяется буквой на N позиций дальше в алфавите.\r\n" +
                    "Поддерживаются латинский и русский алфавиты.\r\n" +
                    "Цифры и знаки препинания не изменяются.\r\n\r\n" +
                    "Пример (сдвиг 3): A -> D, B -> E, Hello -> Khoor\r\n\r\n\r\n" 
                    
            };

            tab3.Controls.AddRange(new Control[] { HelpTitle, rtbHelp });

            //  СОБЫТИЯ 

            tab2.DragEnter += (s, e) =>
            {
                if (e.Data!.GetDataPresent(DataFormats.FileDrop))
                    e.Effect = DragDropEffects.Copy;
            };
            tab2.DragDrop += (s, e) => HandleFileDrop(e);
            lstBatchFiles.DragEnter += (s, e) =>
            {
                if (e.Data!.GetDataPresent(DataFormats.FileDrop))
                    e.Effect = DragDropEffects.Copy;
            };
            lstBatchFiles.DragDrop += (s, e) => HandleFileDrop(e);

            // Зашифровать
            btnEncrypt.Click += (s, e) =>
            {
                if (txtInput == null || string.IsNullOrWhiteSpace(txtInput.Text))
                {
                    MessageBox.Show("Введите текст для шифрования!", "Ошибка",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                if (numShift != null && numShift.Value == 0)
                {
                    MessageBox.Show("Сдвиг равен 0 — текст останется без изменений.", "Предупреждение",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                try
                {
                    if (txtOutput != null && numShift != null)
                        txtOutput.Text = CaesarCipher.Apply(txtInput.Text, (int)numShift.Value);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка шифрования: " + ex.Message, "Ошибка",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };

            // Расшифровать
            btnDecrypt.Click += (s, e) =>
            {
                if (txtInput == null || string.IsNullOrWhiteSpace(txtInput.Text))
                {
                    MessageBox.Show("Введите текст для расшифровки!", "Ошибка",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                if (numShift != null && numShift.Value == 0)
                {
                    MessageBox.Show("Сдвиг равен 0 — текст останется без изменений.", "Предупреждение",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                try
                {
                    if (txtOutput != null && numShift != null)
                        txtOutput.Text = CaesarCipher.Apply(txtInput.Text, -(int)numShift.Value);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка расшифровки: " + ex.Message, "Ошибка",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };

            // Загрузить файл
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
                    MessageBox.Show("Ошибка загрузки файла: " + ex.Message, "Ошибка",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };

            // Сохранить файл
            btnFileSave.Click += (s, e) =>
            {
                if (txtOutput == null || string.IsNullOrEmpty(txtOutput.Text))
                {
                    MessageBox.Show("Нет текста для сохранения!", "Ошибка",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                    MessageBox.Show("Ошибка сохранения: " + ex.Message, "Ошибка",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };

            // Копировать в буфер
            btnCopy.Click += (s, e) =>
            {
                if (txtOutput == null || string.IsNullOrEmpty(txtOutput.Text))
                {
                    MessageBox.Show("Нет текста для копирования!", "Ошибка",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                    MessageBox.Show("Ошибка копирования: " + ex.Message, "Ошибка",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };

            // Добавить файлы в батч
            btnAddBatch.Click += (s, e) =>
            {
                try
                {
                    OpenFileDialog ofd = new OpenFileDialog { Filter = fileFilter, Multiselect = true };
                    if (ofd.ShowDialog() == DialogResult.OK && lstBatchFiles != null)
                        foreach (string f in ofd.FileNames)
                            if (!lstBatchFiles.Items.Contains(f))
                                lstBatchFiles.Items.Add(f);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка добавления файлов: " + ex.Message, "Ошибка",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };

            // Очистить список
            btnClearBatch.Click += (s, e) => lstBatchFiles?.Items.Clear();

            // Обработать батч
            btnProcessBatch.Click += (s, e) =>
            {
                if (lstBatchFiles == null || lstBatchFiles.Items.Count == 0 || numBatchShift == null)
                {
                    MessageBox.Show("Выберите файлы для обработки!", "Ошибка",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var confirm = MessageBox.Show(
                    $"Файлы ({lstBatchFiles.Items.Count} шт.) будут перезаписаны. Продолжить?",
                    "Подтверждение",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);

                if (confirm != DialogResult.Yes) return;

                int successCount = 0;
                List<string> errors = new List<string>();

                int shift = (int)numBatchShift.Value;
                if (cmbBatchMode != null && cmbBatchMode.SelectedIndex == 1)
                    shift = -shift;

                foreach (string path in lstBatchFiles.Items)
                {
                    try
                    {
                        string content = SmartRead(path);
                        SmartSave(path, CaesarCipher.Apply(content, shift));
                        successCount++;
                    }
                    catch (Exception ex)
                    {
                        errors.Add($"{Path.GetFileName(path)}: {ex.Message}");
                    }
                }

                if (errors.Count == 0)
                {
                    MessageBox.Show($"Готово! Все файлы обработаны (сдвиг: {shift}).",
                        "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    string errorList = string.Join("\n", errors);
                    MessageBox.Show(
                        $"Обработано успешно: {successCount}\nОшибки ({errors.Count}):\n{errorList}",
                        "Частичная ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                lstBatchFiles.Items.Clear();
            };

            // Очистить поля
            btnClear.Click += (s, e) =>
            {
                if (txtInput != null) txtInput.Text = "";
                if (txtOutput != null) txtOutput.Text = "";
            };

// ══════════════════════════════════════════════════════════════════════
// ПОЛЯ КЛАССА — добавить рядом с остальными полями вверху Form1
// ══════════════════════════════════════════════════════════════════════
// private TabPage?       tab4;
// private Panel?         eduPanel;
// private Label?         lblEduResult;
// private TextBox?       txtEduInput;
// private NumericUpDown? numEduShift;
// private ListBox?       lstEduSteps;
// ══════════════════════════════════════════════════════════════════════
// ВКЛАДКА 4 — ОБУЧЕНИЕ
// Вставить в конец InitializeCustomComponents(), перед закрывающей }
// ══════════════════════════════════════════════════════════════════════

tab4 = new TabPage { Text = "🎓 ОБУЧЕНИЕ", BackColor = Color.FromArgb(25, 25, 25) };
tabs.TabPages.Add(tab4);

eduPanel = new Panel
{
    Dock       = DockStyle.Fill,
    AutoScroll = true,
    BackColor  = Color.FromArgb(25, 25, 25)
};
tab4.Controls.Add(eduPanel);

const int PX = 15;
const int PW = 360;
int cy = 15;

// хелпер — заголовок секции
void AddSectionHeader(string text)
{
    eduPanel.Controls.Add(new Label
    {
        Text      = text,
        ForeColor = Color.FromArgb(0, 200, 150),
        Font      = new Font("Segoe UI", 10, FontStyle.Bold),
        Location  = new Point(PX, cy),
        Size      = new Size(PW, 20),
        AutoSize  = false
    });
    cy += 26;
}

// хелпер — текст. Высота считается автоматически по количеству строк
void AddText(string text, Color? color = null, Font? font = null)
{
    var f = font ?? new Font("Segoe UI", 9);
    // считаем нужную высоту: кол-во строк * высота строки + запас
    int lines  = text.Split('\n').Length;
    int height = lines * (f.Height + 2) + 8;

    eduPanel.Controls.Add(new Label
    {
        Text      = text,
        ForeColor = color ?? Color.FromArgb(210, 210, 210),
        Font      = f,
        Location  = new Point(PX, cy),
        Size      = new Size(PW, height),
        AutoSize  = false
    });
    cy += height + 12;
}

void AddDivider()
{
    eduPanel.Controls.Add(new Panel
    {
        Location  = new Point(PX, cy),
        Size      = new Size(PW, 1),
        BackColor = Color.FromArgb(55, 55, 70)
    });
    cy += 10;
}

// ════════════════════════════════════════════════════════════════════
// ЗАГОЛОВОК
// ════════════════════════════════════════════════════════════════════
eduPanel.Controls.Add(new Label
{
    Text      = "КАК РАБОТАЕТ ШИФР ЦЕЗАРЯ",
    ForeColor = Color.FromArgb(255, 200, 0),
    Font      = new Font("Segoe UI", 14, FontStyle.Bold),
    Location  = new Point(PX, cy),
    Size      = new Size(PW, 28),
    AutoSize  = false
});
cy += 40;

// ════════════════════════════════════════════════════════════════════
// 1. ЧТО ТАКОЕ ШИФР ЦЕЗАРЯ
// ════════════════════════════════════════════════════════════════════
AddSectionHeader("1. ЧТО ТАКОЕ ШИФР ЦЕЗАРЯ?");
AddText(
    "Шифр Цезаря — один из древнейших методов шифрования,\r\n" +
    "которым пользовался сам Юлий Цезарь.\r\n" +
    "Принцип прост: каждая буква заменяется другой, сдвинутой на N позиций по алфавиту.\r\n" +
    "N называется ключом (сдвигом).\r\n\r\n");
AddDivider();

// ════════════════════════════════════════════════════════════════════
// 2. АЛФАВИТ И СДВИГ
// ════════════════════════════════════════════════════════════════════
AddSectionHeader("2. АЛФАВИТ И СДВИГ — НАГЛЯДНО");

eduPanel.Controls.Add(new Label { Text = "Ориг:",  ForeColor = Color.Gray,                  Font = new Font("Consolas", 8), Location = new Point(PX,      cy), AutoSize = true });
eduPanel.Controls.Add(new Label { Text = "+3:",    ForeColor = Color.FromArgb(0, 150, 255), Font = new Font("Consolas", 8), Location = new Point(PX,      cy + 18), AutoSize = true });

const string ABC = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
const int    CW  = 13;
int          lx  = PX + 38;

for (int i = 0; i < 26; i++)
{
    eduPanel.Controls.Add(new Label { Text = ABC[i].ToString(),          ForeColor = Color.FromArgb(150,150,150), Font = new Font("Consolas", 8),             Size = new Size(CW,16), Location = new Point(lx + i*CW, cy),    TextAlign = ContentAlignment.MiddleCenter });
    eduPanel.Controls.Add(new Label { Text = ABC[(i+3)%26].ToString(),   ForeColor = Color.FromArgb(0,150,255),   Font = new Font("Consolas", 8, FontStyle.Bold), Size = new Size(CW,16), Location = new Point(lx + i*CW, cy+18), TextAlign = ContentAlignment.MiddleCenter });
}
cy += 40;

eduPanel.Controls.Add(new Label
{
    Text = "A→D,  B→E,  …  X→A,  Y→B,  Z→C   (алфавит «закольцован»)",
    ForeColor = Color.FromArgb(110,110,130), Font = new Font("Segoe UI", 8, FontStyle.Italic),
    Location = new Point(PX, cy), Size = new Size(PW, 16), AutoSize = false
});
cy += 24;
AddDivider();

// ════════════════════════════════════════════════════════════════════
// 3. ПОШАГОВЫЙ РАЗБОР
// ════════════════════════════════════════════════════════════════════
AddSectionHeader("3. ПОШАГОВЫЙ РАЗБОР — ПОПРОБУЙ САМ");

eduPanel.Controls.Add(new Label { Text = "Введи слово (латиница) и выбери сдвиг:", ForeColor = Color.Gray, Font = new Font("Segoe UI", 9), Location = new Point(PX, cy), AutoSize = true });
cy += 22;

txtEduInput = new TextBox
{
    Location = new Point(PX, cy), Size = new Size(150, 26),
    BackColor = Color.FromArgb(40,40,52), ForeColor = Color.White,
    BorderStyle = BorderStyle.FixedSingle, Font = new Font("Consolas", 11), Text = "HELLO"
};
numEduShift = new NumericUpDown
{
    Location = new Point(PX+157, cy), Width = 52, Minimum = 1, Maximum = 25, Value = 3,
    BackColor = Color.FromArgb(40,40,52), ForeColor = Color.White,
    BorderStyle = BorderStyle.FixedSingle, Font = new Font("Consolas", 11),
    TextAlign = HorizontalAlignment.Center
};
var btnEduGo = new Button
{
    Text = "▶", Location = new Point(PX+216, cy), Size = new Size(36, 26),
    FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(0,100,200),
    ForeColor = Color.White, Font = new Font("Segoe UI", 10, FontStyle.Bold), Cursor = Cursors.Hand
};
btnEduGo.FlatAppearance.BorderSize = 0;
eduPanel.Controls.AddRange(new Control[] { txtEduInput, numEduShift, btnEduGo });
cy += 34;

lstEduSteps = new ListBox
{
    Location = new Point(PX, cy), Size = new Size(PW, 108),
    BackColor = Color.FromArgb(30,30,40), ForeColor = Color.FromArgb(170,220,170),
    BorderStyle = BorderStyle.FixedSingle, Font = new Font("Consolas", 9),
    SelectionMode = SelectionMode.None
};
eduPanel.Controls.Add(lstEduSteps);
cy += 114;

lblEduResult = new Label
{
    Text = "", ForeColor = Color.FromArgb(0,220,100),
    Font = new Font("Consolas", 11, FontStyle.Bold),
    Location = new Point(PX, cy), Size = new Size(PW, 22), AutoSize = false
};
eduPanel.Controls.Add(lblEduResult);
cy += 30;

btnEduGo.Click += (s, e) =>
{
    if (txtEduInput == null || numEduShift == null || lstEduSteps == null || lblEduResult == null) return;
    string word  = txtEduInput.Text.ToUpper().Trim();
    int    shift = (int)numEduShift.Value;
    lstEduSteps.Items.Clear();
    lblEduResult.Text = "";
    if (string.IsNullOrEmpty(word)) { lstEduSteps.Items.Add("  ⚠  Введи слово!"); return; }
    var res = new StringBuilder();
    foreach (char ch in word)
    {
        if (ch >= 'A' && ch <= 'Z')
        {
            int op = ch-'A', np = (op+shift)%26;
            char nc = (char)('A'+np);
            lstEduSteps.Items.Add($"  '{ch}' (поз.{op,2}) +{shift} → поз.{np,2} → '{nc}'");
            res.Append(nc);
        }
        else { lstEduSteps.Items.Add($"  '{ch}'  →  не буква, без изменений"); res.Append(ch); }
    }
    lblEduResult.Text = $"  {word}  →  {res}";
};
AddDivider();

// ════════════════════════════════════════════════════════════════════
// 4. МАТЕМАТИКА
// ════════════════════════════════════════════════════════════════════
AddSectionHeader("4. МАТЕМАТИКА ШИФРА");
AddText(
    "E(x) = (x + n) mod 26 — шифрование\r\n" +
    "D(x) = (x - n + 26) mod 26 — дешифрование\r\n\r\n" +
    "x — позиция буквы (A=0, B=1, ... Z=25)\r\n" +
    "n — ключ сдвига\r\n" +
    "mod 26 — закольцовывает алфавит",
    font: new Font("Consolas", 9));
AddDivider();

// ════════════════════════════════════════════════════════════════════
// 5. НАДЁЖНОСТЬ
// ════════════════════════════════════════════════════════════════════
AddSectionHeader("5. НАСКОЛЬКО НАДЁЖЕН ШИФР ЦЕЗАРЯ?");
AddText(
    "⚠  В современном мире шифр Цезаря НЕ надёжен:\r\n\r\n" +
    "•  Всего 25 ключей — перебор занимает секунды.\r\n" +
    "•  Частотный анализ: по частоте букв ключ\r\n" +
    "   легко угадывается без перебора.\r\n\r\n" +
    "✅  Подходит только для обучения криптографии.\r\n" +
    "   Не используй для защиты реальных данных!");
AddDivider();

// ════════════════════════════════════════════════════════════════════
// 6. КАК ПОЛЬЗОВАТЬСЯ
// ════════════════════════════════════════════════════════════════════
AddSectionHeader("6. КАК ПОЛЬЗОВАТЬСЯ ПРОГРАММОЙ");
AddText(
    "📝 Вкладка «ТЕКСТ»:\r\n" +
    "   1. Введи текст или загрузи файл.\r\n" +
    "   2. Установи ключ сдвига.\r\n" +
    "   3. Нажми «ЗАШИФРОВАТЬ» или «РАСШИФРОВАТЬ».\r\n" +
    "   4. Скопируй результат или сохрани в файл.\r\n\r\n" +
    "📁 Вкладка «ФАЙЛЫ»:\r\n" +
    "   1. Добавь файлы TXT/DOCX (или перетащи).\r\n" +
    "   2. Выбери ключ и режим.\r\n" +
    "   3. Нажми «НАЧАТЬ ОБРАБОТКУ».\r\n" +
    "   ⚠  Файлы перезапишутся — сделай копии!\r\n\r\n" +
    "💻 Консоль:\r\n" +
    "   .exe e 3 \"Hello\"  →  зашифровать\r\n" +
    "   .exe d 3 \"Khoor\"  →  расшифровать");
AddDivider();

// ════════════════════════════════════════════════════════════════════
// 7. КОД АЛГОРИТМА
// ════════════════════════════════════════════════════════════════════
AddSectionHeader("7. КОД АЛГОРИТМА (ПОД КАПОТОМ)");
AddText(
    "static string Apply(string text, int shift)\r\n" +
    "{\r\n" +
    "  var sb = new StringBuilder();\r\n" +
    "  foreach (char c in text)\r\n" +
    "  {\r\n" +
    "    if (c >= 'A' && c <= 'Z')\r\n" +
    "    {\r\n" +
    "      int pos = (c-'A' + shift%26 + 26) % 26;\r\n" +
    "      sb.Append((char)('A' + pos));\r\n" +
    "    }\r\n" +
    "    else if (c >= 'a' && c <= 'z')\r\n" +
    "    {\r\n" +
    "      int pos = (c-'a' + shift%26 + 26) % 26;\r\n" +
    "      sb.Append((char)('a' + pos));\r\n" +
    "    }\r\n" +
    "    else sb.Append(c);\r\n" +
    "  }\r\n" +
    "  return sb.ToString();\r\n" +
    "}",
    color: Color.FromArgb(170, 210, 255),
    font:  new Font("Consolas", 9));
AddDivider();

// ════════════════════════════════════════════════════════════════════
// 8. МИНИ-ТЕСТ
// ════════════════════════════════════════════════════════════════════
AddSectionHeader("8. МИНИ-ТЕСТ: ПРОВЕРЬ СЕБЯ");

string[] qTexts   = {
    "Вопрос 1.  E('A') при сдвиге 1?\r\n   a) A       b) B       c) Z",
    "Вопрос 2.  E('Z') при сдвиге 3?\r\n   a) C       b) W       c) Z",
    "Вопрос 3.  Дешифрующий сдвиг для ключа 5?\r\n   a) +5      b) -5      c) +21"
};
string[] qCorrect = { "b", "a", "b" };
string[] qExpl    = {
    "A(0)+1 = B(1)  →  ответ b)",
    "Z(25)+3 = 28 mod 26 = 2 = C  →  ответ a)",
    "Расшифровка = сдвиг с минусом  →  ответ b)"
};

var quizBoxes = new TextBox[3];
for (int qi = 0; qi < 3; qi++)
{
    eduPanel.Controls.Add(new Label
    {
        Text = qTexts[qi], ForeColor = Color.FromArgb(210,210,210),
        Font = new Font("Segoe UI", 9),
        Location = new Point(PX, cy), Size = new Size(PW, 36), AutoSize = false
    });
    cy += 40;

    eduPanel.Controls.Add(new Label { Text = "Твой ответ:", ForeColor = Color.Gray, Font = new Font("Segoe UI", 8), Location = new Point(PX, cy+3), AutoSize = true });
    quizBoxes[qi] = new TextBox
    {
        Location = new Point(PX+92, cy), Width = 36, MaxLength = 1,
        BackColor = Color.FromArgb(40,40,52), ForeColor = Color.White,
        BorderStyle = BorderStyle.FixedSingle, Font = new Font("Consolas", 10),
        TextAlign = HorizontalAlignment.Center
    };
    eduPanel.Controls.Add(quizBoxes[qi]);
    cy += 34;

    if (qi < 2) { eduPanel.Controls.Add(new Panel { Location = new Point(PX, cy), Size = new Size(PW, 1), BackColor = Color.FromArgb(50,50,65) }); cy += 10; }
}
cy += 10;

var btnCheck = new Button
{
    Text = "✔  ПРОВЕРИТЬ ОТВЕТЫ", Location = new Point(PX, cy), Size = new Size(PW, 38),
    FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(35,85,35),
    ForeColor = Color.White, Font = new Font("Segoe UI", 10, FontStyle.Bold), Cursor = Cursors.Hand
};
btnCheck.FlatAppearance.BorderSize = 0;
eduPanel.Controls.Add(btnCheck);
cy += 46;

var lblTestResult = new Label
{
    Text = "", ForeColor = Color.White, Font = new Font("Segoe UI", 9),
    Location = new Point(PX, cy), Size = new Size(PW, 90), AutoSize = false
};
eduPanel.Controls.Add(lblTestResult);
cy += 100;

btnCheck.Click += (s, e) =>
{
    var sb = new StringBuilder(); int score = 0;
    for (int qi = 0; qi < 3; qi++)
    {
        bool ok = quizBoxes[qi].Text.Trim().ToLower() == qCorrect[qi];
        if (ok) score++;
        sb.AppendLine(ok ? $"  ✅  Вопрос {qi+1}: верно!    {qExpl[qi]}" : $"  ❌  Вопрос {qi+1}: неверно.  {qExpl[qi]}");
    }
    sb.Append($"\r\n  Итог: {score} из 3");
    lblTestResult.Text      = sb.ToString();
    lblTestResult.ForeColor = score == 3 ? Color.FromArgb(0,220,100) : Color.FromArgb(255,160,50);
};

eduPanel.Controls.Add(new Panel { Location = new Point(0, cy), Size = new Size(1, 24), BackColor = Color.Transparent });
} // Конец InitializeCustomComponents

} // Конец класса Form1

} // Конец пространства имен CaesarProject

