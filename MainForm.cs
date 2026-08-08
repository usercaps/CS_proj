using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using Newtonsoft.Json;

namespace TitleGen
{
    public partial class MainForm : Form
    {
        private TabControl tabControl;
        private TabPage tabParams1, tabParams2, tabTableEditor;

        // Элементы Страницы 1 (Основная)
        private Panel testsPanel1, inputsPanel1;
        private RadioButton radioTip1, radioPeriod1;
        private ComboBox cmbItemMode1;
        private TextBox txtTemplate1;
        private Button btnGenerate1;
        private Dictionary<string, TextBox> inputs1 = new Dictionary<string, TextBox>();
        private Dictionary<string, CheckBox> testCheckboxes1 = new Dictionary<string, CheckBox>();

        // Элементы Страницы 2 (Леша)
        private Panel testsPanel2, inputsPanel2;
        // Эти переменные нужны для логики, даже если кнопок нет на экране
        private RadioButton radioTip2, radioPeriod2;
        private ComboBox cmbItemMode2;
        private TextBox txtTemplate2;
        private Button btnGenerate2;
        private Dictionary<string, TextBox> inputs2 = new Dictionary<string, TextBox>();
        private Dictionary<string, CheckBox> testCheckboxes2 = new Dictionary<string, CheckBox>();

        // Общие элементы
        private ComboBox cmbTables;
        private DataGridView dgvRows;
        private Button btnAddRow, btnDeleteRow, btnSaveConfig;

        private TemplateConfig currentConfig;
        private string currentConfigPath;
        private TableConfig currentTable;

        private Dictionary<string, List<string>> placeholdersCache = new Dictionary<string, List<string>>();

        private List<TableRow> commonEquipment = new List<TableRow>
        {
            new TableRow { testName = "*", values = new List<string> { "", "Барометр-анероид", "М110", "126", "04.25 - 04.26" } },
            new TableRow { testName = "*", values = new List<string> { "", "Комбинированный прибор ", "Testo 625", "61064548/709", "05.25 - 05.26" } }
        };

        private static readonly Dictionary<string, string> FriendlyNames = new Dictionary<string, string>
        {
            { "Имя_изделия", "Имя изделия" },
            { "Имя_изделия2", "Имя изделия 2" },
            { "Номер_изделия", "Номер изделия" },
            { "Номер_изделия2", "Номер изделия 2" },
            { "Рег_Номер_изделия", "Рег. номер изделия" },
            { "Рег_Номер_изделия2", "Рег. номер изделия 2" },
            { "Дата_начала", "Дата начала испытаний" },
            { "Дата_окончания", "Дата окончания испытаний" },
            { "ТНПА", "ТНПА" },
            { "Номер_протокола", "Номер протокола" },
            { "Дата_протокола", "Дата протокола" },
            { "Номер_приказа", "Номер приказа" },
            { "Дата_приказа", "Дата приказа" },
            { "Номер_Изделия", "Номер изделия" } // Для дублирования
        };

        public MainForm()
        {
            Text = "Генерация протокола";
            Width = 900;
            Height = 650;
            StartPosition = FormStartPosition.CenterScreen;
            AutoScroll = true;

            BuildStaticUI();
        }

        private void BuildStaticUI()
        {
            tabControl = new TabControl { Left = 10, Top = 10, Width = 870, Height = 600 };
            tabControl.SelectedIndexChanged += TabControl_SelectedIndexChanged;

            tabParams1 = new TabPage { Text = "Основная" };
            tabParams2 = new TabPage { Text = "Леша" };
            tabTableEditor = new TabPage { Text = "Редактор таблиц" };

            BuildParamsTab(tabParams1, 1);
            BuildParamsTab(tabParams2, 2);
            BuildTableEditorTab(tabTableEditor);

            tabControl.TabPages.Add(tabParams1);
            tabControl.TabPages.Add(tabParams2);
            tabControl.TabPages.Add(tabTableEditor);
            Controls.Add(tabControl);

            // Инициализация первой страницы
            if (radioTip1 != null) radioTip1.Checked = true;
            UpdateTemplatePath(1);
        }

        private void TabControl_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl.SelectedTab == tabTableEditor)
            {
                PopulateTableDropdown();
                BindTableToGrid();
            }
        }

        private void BuildParamsTab(TabPage page, int pageNum)
        {
            // 1. Панель с чекбоксами (ТОЛЬКО для Страницы 1)
            if (pageNum == 1)
            {
                Panel testsPanel = new Panel
                {
                    Left = 10,
                    Top = 10,
                    Width = 250,
                    Height = 450,
                    BorderStyle = BorderStyle.FixedSingle,
                    AutoScroll = true
                };
                page.Controls.Add(testsPanel);
                testsPanel1 = testsPanel;

                string[] tests = {
                    "Повышенная температура", "Пониженная температура", "Циклы температуры",
                    "Давление рабочее", "Давление предельное",
                    "Повышенная влажность", "Пониженная влажность",
                    "Вибрация", "Удары", "Соляной туман", "Безопасность"
                };

                int y = 10;
                foreach (var test in tests)
                {
                    var cb = new CheckBox { Text = test, Left = 10, Top = y, AutoSize = true };
                    testsPanel.Controls.Add(cb);
                    cb.CheckedChanged += (s, ev) => UpdateRowStatuses(1);
                    testCheckboxes1[test] = cb;
                    y += 25;
                }
            }
            else
            {
                // Для страницы 2 создаем пустую панель, чтобы переменная не была null, но не добавляем на форму
                testsPanel2 = new Panel();
            }

            // 2. Расчет позиции X для центрирования на вкладке "Леша"
            int startX = (pageNum == 2) ? 185 : 280;

            // 3. Радиокнопки (ТОЛЬКО для Страницы 1)
            if (pageNum == 1)
            {
                radioTip1 = new RadioButton { Text = "Типовые", Left = startX, Top = 20, AutoSize = true };
                radioPeriod1 = new RadioButton { Text = "Периодические", Left = startX + 100, Top = 20, AutoSize = true };

                foreach (var rb in new[] { radioTip1, radioPeriod1 })
                {
                    rb.CheckedChanged += (s, e) => UpdateTemplatePath(1);
                    page.Controls.Add(rb);
                }

                // Заглушки для страницы 2, чтобы не было ошибок компиляции
                radioTip2 = new RadioButton();
                radioPeriod2 = new RadioButton();
                radioTip2.Checked = true; // Логически выбираем "Типовые" для пути
            }
            else
            {
                // Инициализация переменных для страницы 2 (визуально их нет)
                radioTip2 = new RadioButton();
                radioPeriod2 = new RadioButton();
                radioTip2.Checked = true; // По умолчанию считаем, что шаблон tipovye
            }

            // 4. Поле пути к шаблону
            TextBox txtTemplate = new TextBox { Left = startX, Top = 60, Width = 500 };

            if (pageNum == 2)
            {
                // Жесткий путь для Леши
                string leshaPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "lesha.docx");
                txtTemplate.Text = leshaPath;
                txtTemplate.ReadOnly = true;
                txtTemplate.Visible = false; // Скрываем поле пути
            }

            if (pageNum == 1) txtTemplate1 = txtTemplate; else txtTemplate2 = txtTemplate;
            page.Controls.Add(txtTemplate);

            // 5. Выбор количества изделий
            var lblItemMode = new Label { Text = "Количество изделий:", Left = startX, Top = 90, AutoSize = true };

            ComboBox cmbItemMode = new ComboBox
            {
                Left = startX + 140,
                Top = 88,
                Width = 120,
                DropDownStyle = ComboBoxStyle.DropDownList
            };

            if (pageNum == 1)
            {
                // Страница 1: 1 или 2 изделия (как было раньше, или можно тоже расширить)
                cmbItemMode.Items.Add("1 изделие");
                cmbItemMode.Items.Add("2 изделия");
                cmbItemMode.SelectedIndex = 0;
            }
            else
            {
                // Страница 2 (Леша): от 1 до 100
                for (int i = 1; i <= 100; i++)
                {
                    cmbItemMode.Items.Add($"{i} изд.");
                }
                cmbItemMode.SelectedIndex = 0; // По умолчанию 1
            }

            if (pageNum == 1) cmbItemMode1 = cmbItemMode; else cmbItemMode2 = cmbItemMode;

            cmbItemMode.SelectedIndexChanged += (s, e) =>
            {
                if (pageNum == 1) UpdateTemplatePath(1);
                else UpdateTemplatePath(2);
            };

            page.Controls.Add(lblItemMode);
            page.Controls.Add(cmbItemMode);

            // 6. Панель ввода данных
            Panel inputsPanel = new Panel
            {
                Left = startX,
                Top = 120,
                Width = 500,
                Height = 280,
                AutoScroll = true,
                BorderStyle = BorderStyle.FixedSingle
            };
            if (pageNum == 1) inputsPanel1 = inputsPanel; else inputsPanel2 = inputsPanel;
            page.Controls.Add(inputsPanel);

            // 7. Кнопка генерации
            Button btnGenerate = new Button
            {
                Text = "Сформировать Протокол",
                Left = startX,
                Top = 420,
                Width = 200
            };

            if (pageNum == 1)
            {
                btnGenerate.Click += (s, e) => btnGenerate_Click(1);
                btnGenerate1 = btnGenerate;
            }
            else
            {
                btnGenerate.Click += (s, e) => btnGenerate_Click(2);
                btnGenerate2 = btnGenerate;
            }

            page.Controls.Add(btnGenerate);

            // Старт
            if (pageNum == 1 && radioTip1 != null)
            {
                radioTip1.Checked = true;
                UpdateTemplatePath(1);
            }
            else if (pageNum == 2)
            {
                UpdateTemplatePath(2);
            }
        }

        private void BuildTableEditorTab(TabPage page)
        {
            var lblTable = new Label { Text = "Выберите таблицу:", Left = 20, Top = 20, AutoSize = true };
            cmbTables = new ComboBox
            {
                Left = 200,
                Top = 18,
                Width = 300,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbTables.SelectedIndexChanged += cmbTables_SelectedIndexChanged;
            page.Controls.Add(lblTable);
            page.Controls.Add(cmbTables);

            dgvRows = new DataGridView
            {
                Left = 20,
                Top = 60,
                Width = 820,
                Height = 400,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                MultiSelect = false
            };
            page.Controls.Add(dgvRows);

            btnAddRow = new Button { Text = "Добавить строку", Left = 20, Top = 480, Width = 150 };
            btnDeleteRow = new Button { Text = "Удалить строку", Left = 180, Top = 480, Width = 150 };
            btnSaveConfig = new Button { Text = "Сохранить параметры", Left = 650, Top = 480, Width = 180 };

            btnAddRow.Click += btnAddRow_Click;
            btnDeleteRow.Click += btnDeleteRow_Click;
            btnSaveConfig.Click += btnSaveConfig_Click;

            page.Controls.AddRange(new Control[] { lblTable, cmbTables, dgvRows, btnAddRow, btnDeleteRow, btnSaveConfig });
        }

        private void UpdateTemplatePath(int pageNum)
        {
            RadioButton radioTip = (pageNum == 1) ? radioTip1 : radioTip2;
            RadioButton radioPeriod = (pageNum == 1) ? radioPeriod1 : radioPeriod2;
            ComboBox cmbItemMode = (pageNum == 1) ? cmbItemMode1 : cmbItemMode2;
            TextBox txtTemplate = (pageNum == 1) ? txtTemplate1 : txtTemplate2;
            Panel inputsPanel = (pageNum == 1) ? inputsPanel1 : inputsPanel2;
            var inputs = (pageNum == 1) ? inputs1 : inputs2;

            if (txtTemplate == null) return;

            string fullPath = "";

            if (pageNum == 2)
            {
                // Для Леши всегда один шаблон
                fullPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "lesha.docx");
            }
            else
            {
                // Для Основной страницы логика выбора
                string templateBase = radioTip.Checked ? "tipovye" : "periodich";

                // Определяем суффикс количества
                string suffix = "_1";
                if (cmbItemMode != null && cmbItemMode.SelectedItem != null)
                {
                    string text = cmbItemMode.SelectedItem.ToString();
                    if (text.Contains("2")) suffix = "_2";
                    // Если вдруг на основной тоже будет список до 100, логика усложнится, 
                    // но пока там только 1 и 2.
                }

                string fileName = $"{templateBase}{suffix}.docx";
                fullPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", fileName);
            }

            txtTemplate.Text = fullPath;

            if (File.Exists(fullPath))
            {
                BuildDynamicForm(fullPath, inputsPanel, inputs);
                LoadConfigForEditor(fullPath);
            }
            else
            {
                if (inputsPanel != null) inputsPanel.Controls.Clear();
                if (inputs != null) inputs.Clear();
                if (pageNum == 1) currentConfig = null;
            }
        }

        private void LoadConfigForEditor(string templatePath)
        {
            if (string.IsNullOrEmpty(templatePath)) return;
            string configPath = Path.Combine(Path.GetDirectoryName(templatePath), "config.json");
            if (!File.Exists(configPath))
            {
                currentConfig = CreateDefaultConfig();
                currentConfigPath = configPath;
                // Создаем конфиг только если его нет, но не пугаем пользователя на каждой загрузке
                if (!File.Exists(configPath)) File.WriteAllText(configPath, JsonConvert.SerializeObject(currentConfig, Formatting.Indented));
            }
            else
            {
                try
                {
                    currentConfig = JsonConvert.DeserializeObject<TemplateConfig>(File.ReadAllText(configPath)) ?? CreateDefaultConfig();
                    currentConfigPath = configPath;
                }
                catch
                {
                    currentConfig = CreateDefaultConfig();
                    currentConfigPath = configPath;
                }
            }
        }

        private TemplateConfig CreateDefaultConfig()
        {
            return new TemplateConfig
            {
                tables = new List<TableConfig>
                {
                    new TableConfig
                    {
                        name = "Программа испытаний",
                        bookmark = "Table_Program",
                        columns = new List<string> { "№", "Наименование объекта испытаний", "ТНПА", "Примечание" },
                        rows = new List<TableRow>
                        {
                            new TableRow { testName = "Повышенная температура", values = new List<string> { "1", "Проверка требований", "4.7.1", "" } }
                        }
                    },
                    new TableConfig
                    {
                        name = "СИ и ИО",
                        bookmark = "Table_Equipment",
                        columns = new List<string> { "№", "Наименование", "Тип", "Номер", "Период" },
                        rows = new List<TableRow>()
                    },
                    new TableConfig
                    {
                        name = "Результаты испытаний",
                        bookmark = "Table_Results",
                        columns = new List<string> { "№", "Наименование", "ТТЗ", "ПМ", "Норма", "Факт", "Вывод" },
                        rows = new List<TableRow>()
                    }
                }
            };
        }

        private void PopulateTableDropdown()
        {
            if (cmbTables == null) return;
            cmbTables.Items.Clear();
            if (currentConfig?.tables == null) return;
            foreach (var table in currentConfig.tables)
                cmbTables.Items.Add(table.name);
            if (cmbTables.Items.Count > 0)
                cmbTables.SelectedIndex = 0;
        }

        private void cmbTables_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (currentConfig?.tables == null || cmbTables?.SelectedIndex < 0) return;
            currentTable = currentConfig.tables[cmbTables.SelectedIndex];
            BindTableToGrid();
        }

        private void BindTableToGrid()
        {
            if (dgvRows == null || currentTable == null) return;
            var activeChecks = testCheckboxes1; // Берем чеки первой страницы для редактора

            dgvRows.Columns.Clear();
            dgvRows.Rows.Clear();

            dgvRows.Columns.Add("testName", "Привязка");
            dgvRows.Columns.Add("status", "Статус");

            for (int i = 0; i < (currentTable.columns?.Count ?? 0); i++)
                dgvRows.Columns.Add($"col{i}", currentTable.columns[i]);

            foreach (var row in currentTable.rows)
            {
                var values = new List<string> { row.testName };
                string status = "Активно";
                if (activeChecks.TryGetValue(row.testName, out CheckBox cb) && !cb.Checked)
                    status = "Скрыто";
                values.Add(status);
                values.AddRange(row.values ?? new List<string>());
                int rIdx = dgvRows.Rows.Add(values.ToArray());
                if (status == "Скрыто")
                {
                    dgvRows.Rows[rIdx].DefaultCellStyle.BackColor = Color.LightGray;
                    dgvRows.Rows[rIdx].DefaultCellStyle.ForeColor = Color.Gray;
                }
            }
            SetupTestNameComboBoxColumn();
        }

        private void SetupTestNameComboBoxColumn()
        {
            if (dgvRows == null) return;
            if (dgvRows.Columns["testName"] is DataGridViewComboBoxColumn) return;

            var col = new DataGridViewComboBoxColumn { Name = "testName", HeaderText = "Привязка" };
            col.Items.Add("");
            foreach (var k in testCheckboxes1.Keys) col.Items.Add(k);

            int idx = dgvRows.Columns["testName"].Index;
            dgvRows.Columns.RemoveAt(idx);
            dgvRows.Columns.Insert(idx, col);
        }

        private void UpdateRowStatuses(int pageNum)
        {
            if (dgvRows == null) return;
            var checks = (pageNum == 1) ? testCheckboxes1 : testCheckboxes2;
            if (dgvRows.Columns["status"] == null) return;

            foreach (DataGridViewRow row in dgvRows.Rows)
            {
                if (row.IsNewRow) continue;
                string tn = row.Cells["testName"].Value?.ToString() ?? "";
                string st = "Активно";
                if (checks.TryGetValue(tn, out CheckBox cb) && !cb.Checked) st = "Скрыто";
                row.Cells["status"].Value = st;
                row.DefaultCellStyle.BackColor = (st == "Скрыто") ? Color.LightGray : dgvRows.DefaultCellStyle.BackColor;
                row.DefaultCellStyle.ForeColor = (st == "Скрыто") ? Color.Gray : dgvRows.DefaultCellStyle.ForeColor;
            }
        }

        private void btnAddRow_Click(object sender, EventArgs e)
        {
            if (currentTable == null) return;
            dgvRows.Rows.Add("", "Активно", Enumerable.Repeat("", currentTable.columns.Count).ToArray());
        }

        private void btnDeleteRow_Click(object sender, EventArgs e)
        {
            if (dgvRows.SelectedRows.Count == 0) return;
            dgvRows.Rows.RemoveAt(dgvRows.SelectedRows[0].Index);
        }

        private void btnSaveConfig_Click(object sender, EventArgs e)
        {
            if (currentConfig == null || currentTable == null) return;
            currentTable.rows.Clear();
            foreach (DataGridViewRow row in dgvRows.Rows)
            {
                if (row.IsNewRow) continue;
                var vals = new List<string>();
                for (int i = 2; i < row.Cells.Count; i++) vals.Add(row.Cells[i].Value?.ToString() ?? "");
                currentTable.rows.Add(new TableRow { testName = row.Cells["testName"].Value?.ToString() ?? "", values = vals });
            }
            try
            {
                File.WriteAllText(currentConfigPath, JsonConvert.SerializeObject(currentConfig, Formatting.Indented));
                MessageBox.Show("Сохранено!");
            }
            catch (Exception ex) { MessageBox.Show("Ошибка: " + ex.Message); }
        }

        private void BuildDynamicForm(string path, Panel panel, Dictionary<string, TextBox> inputs)
        {
            if (panel == null || inputs == null) return;
            panel.Controls.Clear();
            inputs.Clear();
            var phs = ExtractPlaceholders(path);
            int y = 10;

            // Проверяем количество изделий для фильтрации _2
            bool isTwoItems = false;
            // Пытаемся определить контекст (грубо)
            if (panel == inputsPanel1 && cmbItemMode1?.SelectedIndex == 1) isTwoItems = true;
            if (panel == inputsPanel2 && cmbItemMode2?.SelectedIndex > 0) isTwoItems = true; // Если выбрано > 1 (индекс 0 это 1 изд)

            foreach (var ph in phs)
            {
                if (!isTwoItems && (ph.Contains("_2") || ph.EndsWith("2"))) continue;
                string txt = FriendlyNames.ContainsKey(ph) ? FriendlyNames[ph] : ph.Replace("_2", " 2");
                panel.Controls.Add(new Label { Text = txt, Left = 10, Top = y + 3, Width = 200 });
                var tb = new TextBox { Left = 220, Top = y, Width = 250 };
                panel.Controls.Add(tb);
                inputs[ph] = tb;
                y += 30;
            }
        }

        private List<string> ExtractPlaceholders(string path)
        {
            if (placeholdersCache.TryGetValue(path, out var c)) return c;
            var list = new List<string>();
            Word.Application app = null; Word.Document doc = null;
            try
            {
                app = new Word.Application();
                doc = app.Documents.Open(path, ReadOnly: true, Visible: false);
                var m = Regex.Matches(doc.Content.Text, @"\{\{([А-Яа-яA-Za-z0-9_]+)\}\}");
                foreach (Match match in m) if (!list.Contains(match.Groups[1].Value)) list.Add(match.Groups[1].Value);
            }
            finally
            {
                if (doc != null) { doc.Close(false); Marshal.ReleaseComObject(doc); }
                if (app != null) { app.Quit(); Marshal.ReleaseComObject(app); }
                GC.Collect(); GC.WaitForPendingFinalizers();
            }
            placeholdersCache[path] = list;
            return list;
        }

        private void btnGenerate_Click(int pageNum)
        {
            RadioButton radioTip = (pageNum == 1) ? radioTip1 : radioTip2;
            RadioButton radioPeriod = (pageNum == 1) ? radioPeriod1 : radioPeriod2;
            ComboBox cmbItemMode = (pageNum == 1) ? cmbItemMode1 : cmbItemMode2;
            TextBox txtTemplate = (pageNum == 1) ? txtTemplate1 : txtTemplate2;
            var inputs = (pageNum == 1) ? inputs1 : inputs2;
            var checks = (pageNum == 1) ? testCheckboxes1 : testCheckboxes2;

            if (txtTemplate == null || !File.Exists(txtTemplate.Text))
            {
                MessageBox.Show("Шаблон не найден!");
                return;
            }

            // Получаем количество изделий
            int itemsCount = 1;
            if (cmbItemMode != null && cmbItemMode.SelectedItem != null)
            {
                string s = cmbItemMode.SelectedItem.ToString();
                int.TryParse(Regex.Match(s, @"\d+").Value, out itemsCount);
                if (itemsCount < 1) itemsCount = 1;
            }

            string configPath = Path.Combine(Path.GetDirectoryName(txtTemplate.Text), "config.json");

            using (var sfd = new SaveFileDialog { Filter = "Word (*.docx)|*.docx", FileName = $"Protocol_{pageNum}.docx" })
            {
                if (sfd.ShowDialog() != DialogResult.OK) return;

                Word.Application app = null; Word.Document doc = null;
                try
                {
                    app = new Word.Application();
                    doc = app.Documents.Open(txtTemplate.Text, ReadOnly: false, Visible: false);

                    ReplacePlaceholdersInDocument(doc, inputs);

                    // Таблицы только для основной страницы
                    if (pageNum == 1 && File.Exists(configPath))
                    {
                        var cfg = JsonConvert.DeserializeObject<TemplateConfig>(File.ReadAllText(configPath));
                        ProcessTablesFromConfig(doc, cfg, checks, false);
                    }

                    ReplacePlaceholdersInDocument(doc, inputs);

                    // Дублирование страниц если нужно
                    if (itemsCount > 1 && doc.Bookmarks.Exists("ItemStart") && doc.Bookmarks.Exists("ItemEnd"))
                    {
                        DuplicateItemPages(doc, itemsCount, inputs);
                    }

                    doc.SaveAs2(sfd.FileName);
                    MessageBox.Show($"Готово! Изделий: {itemsCount}");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка: " + ex.Message);
                }
                finally
                {
                    if (doc != null) { doc.Close(false); Marshal.ReleaseComObject(doc); }
                    if (app != null) { app.Quit(); Marshal.ReleaseComObject(app); }
                    GC.Collect(); GC.WaitForPendingFinalizers();
                }
            }
        }

        private void ReplacePlaceholdersInDocument(Word.Document doc, Dictionary<string, TextBox> inputs)
        {
            foreach (var p in inputs)
            {
                var rng = doc.Content;
                rng.Find.Execute(FindText: "{{" + p.Key + "}}", ReplaceWith: p.Value.Text, Replace: Word.WdReplace.wdReplaceAll);
            }
        }

        private void ProcessTablesFromConfig(Word.Document doc, TemplateConfig cfg, Dictionary<string, CheckBox> checks, bool isTwo)
        {
            foreach (var t in cfg.tables)
            {
                if (!doc.Bookmarks.Exists(t.bookmark)) continue;
                try
                {
                    if (t.name == "Результаты испытаний")
                    {
                        var bk = doc.Bookmarks[t.bookmark];
                        var rng = bk.Range; rng.Text = "\n"; rng.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                        var groups = new Dictionary<string, List<TableRow>>();
                        foreach (var r in t.rows)
                        {
                            if (checks.TryGetValue(r.testName, out CheckBox cb) && cb.Checked)
                            {
                                if (!groups.ContainsKey(r.testName)) groups[r.testName] = new List<TableRow>();
                                groups[r.testName].Add(r);
                            }
                        }
                        if (groups.Count == 0) continue;

                        foreach (var g in groups)
                        {
                            var rows = g.Value.Select(r => new TableRow { testName = r.testName, values = (isTwo && r.valuesTwo != null) ? r.valuesTwo : r.values }).ToList();
                            var tbl = doc.Tables.Add(rng, rows.Count, t.columns.Count);
                            for (int i = 0; i < rows.Count; i++)
                                for (int c = 0; c < t.columns.Count; c++)
                                {
                                    string txt = (c < rows[i].values.Count) ? rows[i].values[c] : "";
                                    tbl.Cell(i + 1, c + 1).Range.Text = txt;
                                    tbl.Cell(i + 1, c + 1).Range.Font.Name = "Times New Roman";
                                    tbl.Cell(i + 1, c + 1).Range.Font.Size = 13;
                                }
                            tbl.Borders.Enable = 1;
                            rng = tbl.Range; rng.Collapse(Word.WdCollapseDirection.wdCollapseEnd); rng.Text = "\n"; rng.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                        }
                    }
                    else
                    {
                        var bk = doc.Bookmarks[t.bookmark];
                        var tbl = bk.Range.Tables[1];
                        int startRow = bk.Range.Rows[1].Index;
                        var list = t.rows.Where(r => checks.TryGetValue(r.testName, out CheckBox cb) && cb.Checked).ToList();

                        if (t.name == "СИ и ИО")
                        {
                            string any = checks.FirstOrDefault(x => x.Value.Checked).Key ?? "";
                            foreach (var eq in commonEquipment) list.Add(new TableRow { testName = any, values = eq.values });
                        }

                        for (int i = 0; i < list.Count; i++)
                        {
                            int rIdx = startRow + i;
                            if (rIdx > tbl.Rows.Count) tbl.Rows.Add();
                            for (int c = 0; c < t.columns.Count; c++)
                            {
                                string txt = (c < list[i].values.Count) ? list[i].values[c] : "";
                                tbl.Cell(rIdx, c + 1).Range.Text = txt;
                                tbl.Cell(rIdx, c + 1).Range.Font.Name = "Times New Roman";
                                tbl.Cell(rIdx, c + 1).Range.Font.Size = 13;
                            }
                            tbl.Cell(rIdx, 1).Range.Text = (i + 1).ToString();
                        }
                    }
                }
                catch { }
            }
        }

        private void DuplicateItemPages(Word.Document doc, int totalCount, Dictionary<string, TextBox> inputs)
        {
            Word.Bookmark startBk = doc.Bookmarks["ItemStart"];
            Word.Bookmark endBk = doc.Bookmarks["ItemEnd"];

            Word.Range sampleRange = doc.Range(startBk.Start, endBk.End);
            Word.Range currentPos = doc.Range(endBk.End, endBk.End);

            for (int i = 2; i <= totalCount; i++)
            {
                sampleRange.Copy();
                currentPos.Paste();

                Word.Range newBlock = doc.Range(currentPos.Start, currentPos.Start + sampleRange.Characters.Count);

                // Замена номера изделия
                Word.Find f = newBlock.Find;
                f.Execute(FindText: "{{Номер_Изделия}}", ReplaceWith: i.ToString(), Replace: Word.WdReplace.wdReplaceAll);
                // Можно добавить замену других уникальных полей если нужно

                currentPos.Start = newBlock.End;
                currentPos.End = newBlock.End;
            }

            // Заполняем оригинал первым номером
            Word.Range orig = doc.Range(startBk.Start, endBk.End);
            Word.Find fOrig = orig.Find;
            fOrig.Execute(FindText: "{{Номер_Изделия}}", ReplaceWith: "1", Replace: Word.WdReplace.wdReplaceAll);

            // Удаляем закладки и оригинальный текст (так как копии уже созданы)
            // Внимание: удаляем диапазон, который занимал оригинал
            orig.Delete();
        }
    }
}