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
        private TabPage tabParams, tabTableEditor;
        private Panel testsPanel, inputsPanel;
        private RadioButton radioTip, radioPeriod, radioTest;
        private ComboBox cmbItemMode;
        private TextBox txtTemplate;
        private Button btnGenerate;
        private ComboBox cmbTables;
        private DataGridView dgvRows;
        private Button btnAddRow, btnDeleteRow, btnSaveConfig;

        private Dictionary<string, TextBox> inputs = new Dictionary<string, TextBox>();
        private Dictionary<string, CheckBox> testCheckboxes = new Dictionary<string, CheckBox>();

        private TemplateConfig currentConfig;
        private string currentConfigPath;
        private TableConfig currentTable;

        // Кэш для ускорения загрузки плейсхолдеров
        private Dictionary<string, List<string>> placeholdersCache = new Dictionary<string, List<string>>();

        private List<TableRow> commonEquipment = new List<TableRow>
        {
            new TableRow { testName = "*", values = new List<string> { "", "Барометр-анероид", "М110", "126", "04.25 - 04.26" } },
            new TableRow { testName = "*", values = new List<string> { "", "Комбинированный прибор ", "Testo 625", "61064548/709", "05.25 - 05.26" } }
        };

        // Словарь для красивых названий полей
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
            { "Дата_приказа", "Дата приказа" }
        };

        public MainForm()
        {
            Text = "Генерация протокола";
            Width = 850;
            Height = 600;
            StartPosition = FormStartPosition.CenterScreen;
            AutoScroll = true;
            BuildStaticUI();
        }

        private void BuildStaticUI()
        {
            tabControl = new TabControl { Left = 10, Top = 10, Width = 820, Height = 550 };
            tabParams = new TabPage { Text = "Параметры" };
            tabTableEditor = new TabPage { Text = "Редактор таблиц" };

            BuildParamsTab(tabParams);
            BuildTableEditorTab(tabTableEditor);

            tabControl.TabPages.Add(tabParams);
            tabControl.TabPages.Add(tabTableEditor);
            Controls.Add(tabControl);
        }

        private void BuildParamsTab(TabPage page)
        {
            testsPanel = new Panel
            {
                Left = 10,
                Top = 10,
                Width = 250,
                Height = 450,
                BorderStyle = BorderStyle.FixedSingle,
                AutoScroll = true
            };
            page.Controls.Add(testsPanel);

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
                testCheckboxes[test] = cb;
                cb.CheckedChanged += (s, ev) => UpdateRowStatuses();
                y += 25;
            }

            radioTip = new RadioButton { Text = "Типовые", Left = 280, Top = 20, AutoSize = true };
            radioPeriod = new RadioButton { Text = "Периодические", Left = 380, Top = 20, AutoSize = true };
            radioTest = new RadioButton { Text = "Тест", Left = 520, Top = 20, AutoSize = true };
            txtTemplate = new TextBox { Left = 280, Top = 60, Width = 500 };

            foreach (var rb in new[] { radioTip, radioPeriod, radioTest })
            {
                rb.CheckedChanged += TemplateSelectorChanged;
                page.Controls.Add(rb);
            }
            page.Controls.Add(txtTemplate);

            var lblItemMode = new Label { Text = "Количество изделий:", Left = 280, Top = 90, AutoSize = true };
            cmbItemMode = new ComboBox
            {
                Left = 420,
                Top = 88,
                Width = 120,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbItemMode.Items.Add("1 изделие");
            cmbItemMode.Items.Add("2 изделия");
            cmbItemMode.SelectedIndex = 0;
            page.Controls.Add(lblItemMode);
            page.Controls.Add(cmbItemMode);

            cmbItemMode.SelectedIndexChanged += (s, e) =>
            {
                UpdateTemplatePath();
            };

            inputsPanel = new Panel
            {
                Left = 280,
                Top = 120,
                Width = 500,
                Height = 280,
                AutoScroll = true,
                BorderStyle = BorderStyle.FixedSingle
            };
            page.Controls.Add(inputsPanel);

            btnGenerate = new Button
            {
                Text = "Сформировать Протокол",
                Left = 280,
                Top = 420,
                Width = 200
            };
            btnGenerate.Click += btnGenerate_Click;
            page.Controls.Add(btnGenerate);
        }

        private void BuildTableEditorTab(TabPage page)
        {
            var lblTable = new Label { Text = "Выберите таблицу:", Left = 20, Top = 20, AutoSize = true };
            cmbTables = new ComboBox
            {
                Left = 150,
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
                Width = 780,
                Height = 350,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                MultiSelect = false
            };
            page.Controls.Add(dgvRows);

            btnAddRow = new Button { Text = "Добавить строку", Left = 20, Top = 420, Width = 150 };
            btnDeleteRow = new Button { Text = "Удалить строку", Left = 180, Top = 420, Width = 150 };
            btnSaveConfig = new Button { Text = "Сохранить параметры", Left = 600, Top = 420, Width = 180 };

            btnAddRow.Click += btnAddRow_Click;
            btnDeleteRow.Click += btnDeleteRow_Click;
            btnSaveConfig.Click += btnSaveConfig_Click;

            page.Controls.AddRange(new Control[] { lblTable, cmbTables, dgvRows, btnAddRow, btnDeleteRow, btnSaveConfig });
        }

        private void TemplateSelectorChanged(object sender, EventArgs e)
        {
            UpdateTemplatePath();
        }

        private void UpdateTemplatePath()
        {
            string templateBase;
            if (radioTip.Checked)
                templateBase = "tipovye";
            else if (radioPeriod.Checked)
                templateBase = "periodich";
            else if (radioTest.Checked)
                templateBase = "test";
            else
                return;

            string suffix = (cmbItemMode.SelectedIndex == 1) ? "_2" : "_1";
            string templateFileName = $"{templateBase}{suffix}.docx";
            string baseDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates");
            string fullPath = Path.Combine(baseDir, templateFileName);

            txtTemplate.Text = fullPath;

            if (File.Exists(fullPath))
            {
                BuildDynamicForm(fullPath);
                LoadConfigForEditor(fullPath);
            }
            else
            {
                inputsPanel.Controls.Clear();
                cmbTables.Items.Clear();
            }
        }

        private void LoadConfigForEditor(string templatePath)
        {
            string configPath = Path.Combine(Path.GetDirectoryName(templatePath), "config.json");
            if (!File.Exists(configPath))
            {
                currentConfig = CreateDefaultConfig();
                currentConfigPath = configPath;
                File.WriteAllText(configPath, JsonConvert.SerializeObject(currentConfig, Formatting.Indented));
                MessageBox.Show($"Создан новый config.json:\n{configPath}");
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
            PopulateTableDropdown();
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
                        columns = new List<string>
                        {
                            "№",
                            "Наименование объекта испытаний (показателей, характеристик)",
                            "Наименование ТНПА, устанавливающего метод испытаний",
                            "Примечание"
                        },
                        rows = new List<TableRow>
                        {
                            new TableRow { testName = "Повышенная температура", values = new List<string> { "1", "Проверка требований к воздействию повышенной рабочей и повышенной предельной температуры", "4.7.1", "" } },
                            new TableRow { testName = "Пониженная температура", values = new List<string> { "2", "Проверка требований к воздействию пониженной рабочей и пониженной предельной температуры", "4.7.2", "" } },
                            new TableRow { testName = "Циклы температуры", values = new List<string> { "3", "Проверка требований к изменению температуры окружающей среды", "4.7.9", "" } },
                            new TableRow { testName = "Давление рабочее", values = new List<string> { "", "Проверка требований к воздействию пониженного рабочего, предельного атмосферного давления", "4.7.3, 4.7.4", "" } },
                            new TableRow { testName = "Удары", values = new List<string> { "", "Проверка устойчивости и прочности при воздействии ударных нагрузок", "4.7.11, а), 4.7.11, б), 4.7.12", "" } }
                        }
                    },
                    new TableConfig
                    {
                        name = "СИ и ИО",
                        bookmark = "Table_Equipment",
                        columns = new List<string>
                        {
                            "№",
                            "Наименование испытательного оборудования и средств измерений",
                            "Тип, марка",
                            "Номер",
                            "Период аттестации, калибровки"
                        },
                        rows = new List<TableRow>
                        {
                            new TableRow { testName = "Вибрация", values = new List<string> { "", "Вибростенд LDS V408", "VS-408-001", "2025-11-30", "" } },
                            new TableRow { testName = "Повышенная температура", values = new List<string> { "", "Камера тепла и холода", "МС-71", "906569", "08.24 - 08.25" } },
                            new TableRow { testName = "Пониженная температура", values = new List<string> { "", "Камера тепла и холода", "МС-71", "906569", "08.24 - 08.25" } },
                            new TableRow { testName = "Циклы температуры", values = new List<string> { "", "Камера тепла и холода", "МС-71", "906569", "08.24 - 08.25" } },
                            new TableRow { testName = "Удары", values = new List<string> { "", "Ударная установка", "STT500", "2/79", "10.24 - 10.25" } },
                            new TableRow { testName = "Давление рабочее", values = new List<string> { "", "Термобарокамера", "TBV-2000", "308934", "08.24 - 08.25" } },
                            new TableRow { testName = "Давление предельное", values = new List<string> { "", "Термобарокамера", "TBV-2000", "308934", "08.24 - 08.25" } }
                        }
                    },
                    new TableConfig
                    {
                        name = "Результаты испытаний",
                        bookmark = "Table_Results",
                        columns = new List<string>
                        {
                            "№",
                            "Наименование объекта испытаний (показателей, характеристик)",
                            "ТТЗ (требования)",
                            "ПМ (методы)",
                            "Нормированное значение показателей, установленных в ТНПА",
                            "Фактические значения показателей",
                            "Вывод о соответствии требованиям ТНПА"
                        },
                        rows = new List<TableRow>
                        {
                            new TableRow { testName = "Повышенная температура", values = new List<string> { "1", "Проверка воздействия повышенной температуры", "4.7.1", "ГОСТ Р 57200-2016", "от -60 до +85°C", "+85°C", "Соответствует" } },
                            new TableRow { testName = "Пониженная температура", values = new List<string> { "2", "Проверка воздействия пониженной температуры", "4.7.2", "ГОСТ Р 57200-2016", "от -60 до +85°C", "-60°C", "Соответствует" } },
                            new TableRow { testName = "Циклы температуры", values = new List<string> { "3", "Проверка циклов температуры", "4.7.9", "ГОСТ Р 57200-2016", "10 циклов", "10 циклов", "Соответствует" } },
                            new TableRow { testName = "Давление рабочее", values = new List<string> { "", "Проверка давления", "4.7.3", "ГОСТ Р 57200-2016", "760 мм рт.ст.", "755 мм рт.ст.", "Соответствует" } },
                            new TableRow { testName = "Давление предельное", values = new List<string> { "", "Проверка предельного давления", "4.7.4", "ГОСТ Р 57200-2016", "400 мм рт.ст.", "410 мм рт.ст.", "Соответствует" } },
                            new TableRow { testName = "Удары", values = new List<string> { "", "Проверка ударов", "4.7.11", "ГОСТ Р 57200-2016", "9g, 6 мс", "9g, 6 мс", "Соответствует" } }
                        }
                    }
                }
            };
        }

        private void PopulateTableDropdown()
        {
            cmbTables.Items.Clear();
            if (currentConfig?.tables == null) return;
            foreach (var table in currentConfig.tables)
                cmbTables.Items.Add(table.name);
            if (cmbTables.Items.Count > 0)
                cmbTables.SelectedIndex = 0;
        }

        private void cmbTables_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (currentConfig?.tables == null || cmbTables.SelectedIndex < 0) return;
            currentTable = currentConfig.tables[cmbTables.SelectedIndex];
            BindTableToGrid();
        }

        private void BindTableToGrid()
        {
            dgvRows.Columns.Clear();
            dgvRows.Rows.Clear();
            if (currentTable?.rows == null) return;

            dgvRows.Columns.Add("testName", "Привязка к чекбоксу");
            dgvRows.Columns.Add("status", "Статус");

            for (int i = 0; i < (currentTable.columns?.Count ?? 0); i++)
            {
                string colName = currentTable.columns[i];
                dgvRows.Columns.Add($"col{i}", colName);
            }

            foreach (var row in currentTable.rows)
            {
                var values = new List<string> { row.testName };
                string status = "Активно";
                if (testCheckboxes.TryGetValue(row.testName, out CheckBox cb) && !cb.Checked)
                    status = "Скрыто";
                values.Add(status);
                values.AddRange(row.values ?? new List<string>());
                var rowIndex = dgvRows.Rows.Add(values.ToArray());
                if (status == "Скрыто")
                {
                    dgvRows.Rows[rowIndex].DefaultCellStyle.BackColor = Color.LightGray;
                    dgvRows.Rows[rowIndex].DefaultCellStyle.ForeColor = Color.Gray;
                }
            }

            SetupTestNameComboBoxColumn();
            if (dgvRows.Columns["status"] != null)
            {
                dgvRows.Columns["status"].ReadOnly = true;
                dgvRows.Columns["status"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
        }

        private void SetupTestNameComboBoxColumn()
        {
            if (dgvRows.Columns["testName"] is DataGridViewComboBoxColumn) return;

            var comboBoxColumn = new DataGridViewComboBoxColumn
            {
                Name = "testName",
                HeaderText = "Привязка к чекбоксу"
            };
            comboBoxColumn.Items.Add("");
            foreach (var testName in testCheckboxes.Keys)
                comboBoxColumn.Items.Add(testName);

            int colIndex = dgvRows.Columns["testName"].Index;
            dgvRows.Columns.RemoveAt(colIndex);
            dgvRows.Columns.Insert(colIndex, comboBoxColumn);
        }

        private void UpdateRowStatuses()
        {
            if (dgvRows.Columns["status"] == null) return;
            foreach (DataGridViewRow row in dgvRows.Rows)
            {
                if (row.IsNewRow) continue;
                string testName = row.Cells["testName"].Value?.ToString() ?? "";
                string status = "Активно";
                if (testCheckboxes.TryGetValue(testName, out CheckBox cb) && !cb.Checked)
                    status = "Скрыто";
                row.Cells["status"].Value = status;
                if (status == "Скрыто")
                {
                    row.DefaultCellStyle.BackColor = Color.LightGray;
                    row.DefaultCellStyle.ForeColor = Color.Gray;
                }
                else
                {
                    row.DefaultCellStyle.BackColor = dgvRows.DefaultCellStyle.BackColor;
                    row.DefaultCellStyle.ForeColor = dgvRows.DefaultCellStyle.ForeColor;
                }
            }
        }

        private void btnAddRow_Click(object sender, EventArgs e)
        {
            if (currentTable == null) return;
            int rowIndex = dgvRows.Rows.Add("", "Активно", Enumerable.Repeat("", currentTable.columns.Count).ToArray());
            if (testCheckboxes.Count > 0)
                dgvRows.Rows[rowIndex].Cells["testName"].Value = testCheckboxes.Keys.First();
        }

        private void btnDeleteRow_Click(object sender, EventArgs e)
        {
            if (dgvRows.SelectedRows.Count == 0) return;
            dgvRows.Rows.RemoveAt(dgvRows.SelectedRows[0].Index);
        }

        private void btnSaveConfig_Click(object sender, EventArgs e)
        {
            if (currentConfig == null || currentTable == null)
            {
                MessageBox.Show("Нет данных для сохранения.");
                return;
            }

            currentTable.rows.Clear();
            foreach (DataGridViewRow row in dgvRows.Rows)
            {
                if (row.IsNewRow) continue;

                var values = new List<string>();
                for (int i = 2; i < row.Cells.Count; i++)
                    values.Add(row.Cells[i].Value?.ToString() ?? "");
                currentTable.rows.Add(new TableRow
                {
                    testName = row.Cells["testName"].Value?.ToString() ?? "",
                    values = values
                });
            }

            try
            {
                string json = JsonConvert.SerializeObject(currentConfig, Formatting.Indented);
                File.WriteAllText(currentConfigPath, json);
                MessageBox.Show($"Конфиг сохранён:\n{currentConfigPath}", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                BindTableToGrid();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка сохранения:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BuildDynamicForm(string templatePath)
        {
            inputsPanel.Controls.Clear();
            inputs.Clear();

            var placeholders = ExtractPlaceholders(templatePath);
            int y = 10;

            bool isTwoItems = (cmbItemMode?.SelectedIndex == 1);

            foreach (var ph in placeholders)
            {
                if (!isTwoItems && (ph.Contains("_2") || ph.Contains("2")))
                    continue;

                string displayText = FriendlyNames.ContainsKey(ph)
                    ? FriendlyNames[ph]
                    : ph.Replace("_2", " 2");

                var lbl = new Label { Text = displayText, Left = 10, Top = y + 3, Width = 200 };
                var tb = new TextBox { Left = 220, Top = y, Width = 250 };
                inputsPanel.Controls.Add(lbl);
                inputsPanel.Controls.Add(tb);
                inputs[ph] = tb;
                y += 30;
            }
        }

        // === КЭШИРОВАНИЕ ПЛЕЙСХОЛДЕРОВ ===
        private List<string> ExtractPlaceholders(string path)
        {
            if (placeholdersCache.TryGetValue(path, out List<string> cached))
                return cached;

            var placeholders = new List<string>();
            Word.Application wordApp = null;
            Word.Document doc = null;
            try
            {
                wordApp = new Word.Application();
                doc = wordApp.Documents.Open(path, ReadOnly: true, Visible: false);
                var text = doc.Content.Text;
                var matches = Regex.Matches(text, @"\{\{([А-Яа-яA-Za-z0-9_]+)\}\}");
                foreach (Match match in matches)
                {
                    string ph = match.Groups[1].Value;
                    if (!placeholders.Contains(ph))
                        placeholders.Add(ph);
                }
            }
            finally
            {
                if (doc != null) { doc.Close(false); Marshal.ReleaseComObject(doc); }
                if (wordApp != null) { wordApp.Quit(); Marshal.ReleaseComObject(wordApp); }
                GC.Collect(); GC.WaitForPendingFinalizers();
            }

            placeholdersCache[path] = placeholders;
            return placeholders;
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            if (!File.Exists(txtTemplate.Text))
            {
                MessageBox.Show("Шаблон не найден!");
                return;
            }

            string configPath = Path.Combine(Path.GetDirectoryName(txtTemplate.Text), "config.json");
            if (!File.Exists(configPath))
            {
                MessageBox.Show("Конфиг config.json не найден!");
                return;
            }

            using (var sfd = new SaveFileDialog { Filter = "Word Document (*.docx)|*.docx", FileName = "Протокол.docx" })
            {
                if (sfd.ShowDialog() != DialogResult.OK) return;

                Word.Application wordApp = null;
                Word.Document doc = null;
                try
                {
                    wordApp = new Word.Application();
                    doc = wordApp.Documents.Open(txtTemplate.Text, ReadOnly: false, Visible: false);

                    ReplacePlaceholdersInDocument(doc);

                    var config = currentConfig;
                    ProcessTablesFromConfig(doc, config);
                    ReplacePlaceholdersInDocument(doc);

                    doc.SaveAs2(sfd.FileName);
                    MessageBox.Show("Протокол успешно создан:\n" + sfd.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка при генерации:\n" + ex.Message);
                }
                finally
                {
                    if (doc != null) { doc.Close(false); Marshal.ReleaseComObject(doc); }
                    if (wordApp != null) { wordApp.Quit(); Marshal.ReleaseComObject(wordApp); }
                    GC.Collect(); GC.WaitForPendingFinalizers();
                }
            }
        }

        private void ReplacePlaceholdersInDocument(Word.Document doc)
        {
            foreach (var pair in inputs)
            {
                string placeholder = "{{" + pair.Key + "}}";
                string value = pair.Value.Text;
                var range = doc.Content;
                range.Find.Execute(FindText: placeholder, ReplaceWith: value, Replace: Word.WdReplace.wdReplaceAll);
            }
        }

        private void ProcessTablesFromConfig(Word.Document doc, TemplateConfig config)
        {
            foreach (var tableConfig in config.tables)
            {
                if (!doc.Bookmarks.Exists(tableConfig.bookmark))
                {
                    MessageBox.Show($"Закладка '{tableConfig.bookmark}' не найдена для таблицы '{tableConfig.name}'.");
                    continue;
                }

                try
                {
                    if (tableConfig.name == "Результаты испытаний")
                    {
                        Word.Bookmark bookmark = doc.Bookmarks[tableConfig.bookmark];
                        var range = bookmark.Range;

                        // Очищаем закладку, оставляя абзац
                        range.Text = "\n";
                        range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                        var groups = new Dictionary<string, List<TableRow>>();
                        foreach (var row in tableConfig.rows)
                        {
                            if (testCheckboxes.TryGetValue(row.testName, out CheckBox cb) && cb.Checked)
                            {
                                if (!groups.ContainsKey(row.testName))
                                    groups[row.testName] = new List<TableRow>();
                                groups[row.testName].Add(row);
                            }
                        }

                        if (groups.Count == 0)
                        {
                            MessageBox.Show("Нет выбранных испытаний для результатов.", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            continue;
                        }

                        bool isTwoItems = (cmbItemMode?.SelectedIndex == 1);

                        foreach (var group in groups)
                        {
                            var rowsInGroup = new List<TableRow>();
                            foreach (var row in group.Value)
                            {
                                var selectedValues = (isTwoItems && row.valuesTwo != null)
                                    ? row.valuesTwo
                                    : row.values;

                                rowsInGroup.Add(new TableRow
                                {
                                    testName = row.testName,
                                    values = selectedValues ?? new List<string>()
                                });
                            }

                            Word.Table newTable = doc.Tables.Add(
                                Range: range,
                                NumRows: rowsInGroup.Count,
                                NumColumns: tableConfig.columns.Count
                            );

                            for (int i = 0; i < rowsInGroup.Count; i++)
                            {
                                var rowData = rowsInGroup[i];
                                for (int c = 0; c < tableConfig.columns.Count; c++)
                                {
                                    string text = (c < rowData.values.Count) ? rowData.values[c] : "";
                                    newTable.Cell(i + 1, c + 1).Range.Text = text;

                                    // Шрифт
                                    newTable.Cell(i + 1, c + 1).Range.Font.Name = "Times New Roman";
                                    newTable.Cell(i + 1, c + 1).Range.Font.Size = 13;
                                    newTable.Cell(i + 1, c + 1).Range.Font.Color = Word.WdColor.wdColorBlack;
                                    newTable.Cell(i + 1, c + 1).Shading.BackgroundPatternColor = Word.WdColor.wdColorWhite;

                                    // Выравнивание
                                    newTable.Cell(i + 1, c + 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                                }
                            }

                            // Границы
                            newTable.Borders.Enable = 1;
                            newTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                            newTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                            newTable.Borders.OutsideColor = Word.WdColor.wdColorBlack;
                            newTable.Borders.InsideColor = Word.WdColor.wdColorBlack;

                            // Фиксированные ширины для "Результаты испытаний"
                            if (newTable.Columns.Count >= 7)
                            {
                                newTable.Columns[1].Width = 30;   // №
                                newTable.Columns[2].Width = 200;  // Наименование объекта
                                newTable.Columns[3].Width = 80;   // ТТЗ
                                newTable.Columns[4].Width = 80;   // ПМ
                                newTable.Columns[5].Width = 120;  // Нормированное значение
                                newTable.Columns[6].Width = 120;  // Фактические значения
                                newTable.Columns[7].Width = 80;   // Вывод
                            }

                            // Добавляем пробел после таблицы
                            range = newTable.Range;
                            range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                            range.Text = "\n";
                            range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                        }
                    }
                    else
                    {
                        Word.Bookmark bookmark = doc.Bookmarks[tableConfig.bookmark];
                        Word.Table existingTable = bookmark.Range.Tables[1];
                        int insertRowIndex = bookmark.Range.Rows[1].Index;

                        var rowsToInsert = new List<TableRow>();
                        foreach (var row in tableConfig.rows)
                        {
                            if (testCheckboxes.TryGetValue(row.testName, out CheckBox cb) && cb.Checked)
                                rowsToInsert.Add(row);
                        }

                        if (tableConfig.name == "СИ и ИО")
                        {
                            string anyTest = "";
                            foreach (var kvp in testCheckboxes)
                            {
                                if (kvp.Value.Checked)
                                {
                                    anyTest = kvp.Key;
                                    break;
                                }
                            }
                            foreach (var eq in commonEquipment)
                            {
                                rowsToInsert.Add(new TableRow { testName = anyTest, values = new List<string>(eq.values) });
                            }
                        }

                        if (rowsToInsert.Count == 0)
                        {
                            MessageBox.Show($"Нет выбранных испытаний для таблицы '{tableConfig.name}'.", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            continue;
                        }

                        for (int i = 0; i < rowsToInsert.Count; i++)
                        {
                            int currentRow = insertRowIndex + i;
                            if (currentRow > existingTable.Rows.Count)
                                existingTable.Rows.Add();

                            var rowData = rowsToInsert[i];
                            for (int c = 0; c < tableConfig.columns.Count; c++)
                            {
                                string text = (c < rowData.values.Count) ? rowData.values[c] : "";
                                existingTable.Cell(currentRow, c + 1).Range.Text = text;
                                existingTable.Cell(currentRow, c + 1).Range.Font.Name = "Times New Roman";
                                existingTable.Cell(currentRow, c + 1).Range.Font.Size = 13;
                            }

                            existingTable.Cell(currentRow, 1).Range.Text = (i + 1).ToString();
                        }

                        // Настройка ширины для других таблиц
                        if (tableConfig.name == "Программа испытаний" && existingTable.Columns.Count >= 4)
                        {
                            existingTable.Columns[1].Width = 30;
                            existingTable.Columns[2].Width = 250;
                            existingTable.Columns[3].Width = 150;
                            existingTable.Columns[4].Width = 80;
                        }
                        else if (tableConfig.name == "СИ и ИО" && existingTable.Columns.Count >= 5)
                        {
                            existingTable.Columns[1].Width = 30;
                            existingTable.Columns[2].Width = 200;
                            existingTable.Columns[3].Width = 100;
                            existingTable.Columns[4].Width = 80;
                            existingTable.Columns[5].Width = 100;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"❌ Ошибка вставки данных в '{tableConfig.name}': {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }
}