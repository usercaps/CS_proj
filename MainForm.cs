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

        private Panel testsPanel;
        private Panel inputsPanel;
        private RadioButton radioTip, radioPeriod, radioTest;
        private TextBox txtTemplate;
        private Button btnGenerate;

        private Dictionary<string, TextBox> inputs = new Dictionary<string, TextBox>();
        private Dictionary<string, CheckBox> testCheckboxes = new Dictionary<string, CheckBox>();

        private ComboBox cmbTables;
        private DataGridView dgvRows;
        private Button btnAddRow, btnDeleteRow, btnSaveConfig;
        private TemplateConfig currentConfig;
        private string currentConfigPath;
        private TableConfig currentTable;

        private List<TableRow> commonEquipment = new List<TableRow>
        {
            new TableRow { testName = "*", values = new List<string> { "", "Барометр-анероид", "М110", "126", "04.25 - 04.26" } },
            new TableRow { testName = "*", values = new List<string> { "", "Комбинированный прибор ", "Testo 625", "61064548/709", "05.25 - 05.26" } }
        };

        private Dictionary<string, List<string>> testGroups = new Dictionary<string, List<string>>
        {
            { "Температура", new List<string> { "Повышенная температура", "Пониженная температура", "Циклы температуры" } },
            { "Давление", new List<string> { "Давление рабочее", "Давление предельное" } },
            { "Влажность", new List<string> { "Повышенная влажность", "Пониженная влажность" } }
        };

        public MainForm()
        {
            Text = "Генерация протокола (DocX)";
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
                "Повышенная температура",
                "Пониженная температура",
                "Циклы температуры",
                "Давление рабочее",
                "Давление предельное",
                "Повышенная влажность",
                "Пониженная влажность",
                "Вибрация",
                "Удары",
                "Соляной туман",
                "Безопасность"
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

            inputsPanel = new Panel
            {
                Left = 280,
                Top = 100,
                Width = 500,
                Height = 300,
                AutoScroll = true,
                BorderStyle = BorderStyle.FixedSingle
            };
            page.Controls.Add(inputsPanel);

            btnGenerate = new Button
            {
                Text = "Сформировать DOCX",
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
            page.Controls.Add(lblTable);

            cmbTables = new ComboBox
            {
                Left = 150,
                Top = 18,
                Width = 300,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbTables.SelectedIndexChanged += cmbTables_SelectedIndexChanged;
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
            btnSaveConfig = new Button { Text = "Сохранить config.json", Left = 600, Top = 420, Width = 180 };

            btnAddRow.Click += btnAddRow_Click;
            btnDeleteRow.Click += btnDeleteRow_Click;
            btnSaveConfig.Click += btnSaveConfig_Click;

            page.Controls.AddRange(new Control[] { lblTable, cmbTables, dgvRows, btnAddRow, btnDeleteRow, btnSaveConfig });
        }

        private void TemplateSelectorChanged(object sender, EventArgs e)
        {
            string baseDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates");
            if (radioTip.Checked)
                txtTemplate.Text = Path.Combine(baseDir, "tipovye.docx");
            else if (radioPeriod.Checked)
                txtTemplate.Text = Path.Combine(baseDir, "periodich.docx");
            else if (radioTest.Checked)
                txtTemplate.Text = Path.Combine(baseDir, "test.docx");

            if (File.Exists(txtTemplate.Text))
            {
                BuildDynamicForm(txtTemplate.Text);
                LoadConfigForEditor(txtTemplate.Text);
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
                values.AddRange(row.values);
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
            foreach (var ph in placeholders)
            {
                var lbl = new Label { Text = ph, Left = 10, Top = y + 3, Width = 200 };
                var tb = new TextBox { Left = 220, Top = y, Width = 250 };
                inputsPanel.Controls.Add(lbl);
                inputsPanel.Controls.Add(tb);
                inputs[ph] = tb;
                y += 30;
            }
        }

        private List<string> ExtractPlaceholders(string path)
        {
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

                    // Используем текущий конфиг
                    var config = currentConfig;

                    ProcessTablesFromConfig(doc, config);
                    ReplacePlaceholdersInDocument(doc);

                    doc.SaveAs2(sfd.FileName);
                    MessageBox.Show("DOCX успешно создан:\n" + sfd.FileName);
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

        // ✅ ОСНОВНОЙ МЕТОД С ИЗМЕНЁННОЙ ЛОГИКОЙ ДЛЯ "РЕЗУЛЬТАТЫ ИСПЫТАНИЙ"
        private void ProcessTablesFromConfig(Word.Document doc, TemplateConfig config)
        {
            foreach (var tableConfig in config.tables)
            {
                if (!doc.Bookmarks.Exists(tableConfig.bookmark))
                {
                    MessageBox.Show($"Закладка '{tableConfig.bookmark}' не найдена для таблицы '{tableConfig.name}'.");
                    continue;
                }

                Word.Bookmark bookmark = doc.Bookmarks[tableConfig.bookmark];
                Word.Range insertRange = bookmark.Range;
                insertRange.Text = "";

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
                    MessageBox.Show($"Таблица '{tableConfig.name}' пуста. Пропуск.", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    continue;
                }

                int colCount = tableConfig.columns?.Count ?? 0;
                if (colCount <= 0)
                {
                    MessageBox.Show($"Таблица '{tableConfig.name}' не имеет колонок. Пропуск.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    continue;
                }

                try
                {
                    int totalRows = tableConfig.name == "Результаты испытаний"
                        ? rowsToInsert.Count + 2
                        : rowsToInsert.Count + 1;

                    Word.Table newTable = doc.Tables.Add(
                        insertRange,
                        totalRows,
                        colCount,
                        Word.WdDefaultTableBehavior.wdWord9TableBehavior,
                        Word.WdAutoFitBehavior.wdAutoFitContent
                    );

                    // Границы
                    foreach (Word.Border border in newTable.Borders)
                    {
                        border.LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                        border.LineWidth = Word.WdLineWidth.wdLineWidth050pt;
                        border.Color = Word.WdColor.wdColorAutomatic;
                    }

                    // === СПЕЦИАЛЬНАЯ ОБРАБОТКА ДЛЯ "РЕЗУЛЬТАТЫ ИСПЫТАНИЙ" ===
                    if (tableConfig.name == "Результаты испытаний")
                    {
                        // Убедимся, что таблица имеет 7 колонок
                        if (colCount != 7)
                        {
                            MessageBox.Show("Ошибка: таблица 'Результаты испытаний' должна иметь 7 колонок.");
                            continue;
                        }

                        // --- Заполняем ВСЕ ячейки первых 4 строк ---
                        // Первая строка
                        newTable.Cell(1, 1).Range.Text = "№";
                        newTable.Cell(1, 2).Range.Text = "Наименование объекта испытаний (показателей, характеристик)";
                        newTable.Cell(1, 3).Range.Text = "Номер пункта ТНПА, устанавливающего" + Environment.NewLine + "БФИД 466535.019 ТУ";
                        newTable.Cell(1, 4).Range.Text = ""; // будет объединена с (1,3)
                        newTable.Cell(1, 5).Range.Text = "Нормированное значение показателей, установленных в ТНПА";
                        newTable.Cell(1, 6).Range.Text = "Фактические значения показателей";
                        newTable.Cell(1, 7).Range.Text = "Вывод о соответствии требованиям ТНПА";

                        // Вторая строка — оставляем пустой (будет объединена)
                        for (int c = 1; c <= 7; c++)
                            newTable.Cell(2, c).Range.Text = "";

                        // Третья строка — подзаголовки
                        newTable.Cell(3, 1).Range.Text = "";
                        newTable.Cell(3, 2).Range.Text = "";
                        newTable.Cell(3, 3).Range.Text = "требования";
                        newTable.Cell(3, 4).Range.Text = "методы";
                        newTable.Cell(3, 5).Range.Text = "";
                        newTable.Cell(3, 6).Range.Text = "";
                        newTable.Cell(3, 7).Range.Text = "";

                        // Четвёртая строка — нумерация
                        for (int c = 1; c <= 7; c++)
                            newTable.Cell(4, c).Range.Text = c.ToString();


                        // --- ОБЪЕДИНЕНИЕ ЯЧЕЕК ---
                        // Объединяем (1,3) и (1,4) → горизонтально
                        newTable.Cell(1, 3).Merge(newTable.Cell(1, 4));


                        // 1. Объединяем колонку № (1,1) → (2,1) → (3,1) → вертикально
                        Word.Cell col1 = newTable.Cell(1, 1);
                        col1.Merge(newTable.Cell(2, 1)); // объединили 1 и 2
                        

                        // 2. Объединяем колонку "Наименование..." (1,2) → (2,2) → (3,2)
                        Word.Cell col2 = newTable.Cell(1, 2);
                        col2.Merge(newTable.Cell(2, 2));

                        // 5. Объединяем колонку "Наименование..." (1,5) → (2,5) → (3,5)
                        Word.Cell col5 = newTable.Cell(1, 5);
                        col5.Merge(newTable.Cell(2, 5));

                        // 6. Объединяем колонку "Наименование..." (1,6) → (2,6) → (3,6)
                        Word.Cell col6 = newTable.Cell(1, 6);
                        col6.Merge(newTable.Cell(2, 6));
                        

                        // 7. Объединяем колонку "Наименование..." (1,7) → (2,7) → (3,7)
                        Word.Cell col7 = newTable.Cell(1, 7);
                        col7.Merge(newTable.Cell(2, 7));
                        

                        // --- Заполняем ДАННЫЕ начиная с 5-й строки ---
                        for (int r = 0; r < rowsToInsert.Count; r++)
                        {
                            var rowData = rowsToInsert[r];
                            for (int c = 0; c < colCount; c++)
                            {
                                string cellText = c < rowData.values.Count ? rowData.values[c] : "";
                                newTable.Cell(r + 5, c + 1).Range.Text = cellText;
                            }
                        }

                        // Нумерация строк данных
                        for (int r = 0; r < rowsToInsert.Count; r++)
                        {
                            newTable.Cell(r + 5, 1).Range.Text = (r + 1).ToString();
                        }

                        // --- Форматирование ---
                        for (int r = 1; r <= 4; r++)
                        {
                            for (int c = 1; c <= 7; c++)
                            {
                                Word.Cell cell = newTable.Cell(r, c);
                                cell.Range.Font.Name = "Times New Roman";
                                cell.Range.Font.Size = 13;
                                cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                                cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                                cell.Range.ParagraphFormat.SpaceAfter = 0;
                                cell.Range.ParagraphFormat.SpaceBefore = 0;
                            }
                        }
                    }
                    else
                    {
                        // Обычные таблицы
                        for (int c = 0; c < colCount; c++)
                        {
                            string headerText = c < tableConfig.columns.Count ? tableConfig.columns[c] : "";
                            newTable.Cell(1, c + 1).Range.Text = headerText;
                        }

                        for (int r = 0; r < rowsToInsert.Count; r++)
                        {
                            var rowData = rowsToInsert[r];
                            for (int c = 0; c < colCount; c++)
                            {
                                string cellText = c < rowData.values.Count ? rowData.values[c] : "";
                                newTable.Cell(r + 2, c + 1).Range.Text = cellText;
                            }
                        }

                        // Нумерация для "СИ и ИО"
                        if (tableConfig.name == "СИ и ИО")
                        {
                            for (int r = 0; r < rowsToInsert.Count; r++)
                            {
                                newTable.Cell(r + 2, 1).Range.Text = (r + 1).ToString();
                            }
                        }
                    }

                    // Форматирование ячеек
                    for (int r = 1; r <= newTable.Rows.Count; r++)
                    {
                        for (int c = 1; c <= newTable.Columns.Count; c++)
                        {
                            Word.Cell cell = newTable.Cell(r, c);
                            cell.Range.Font.Name = "Times New Roman";
                            cell.Range.Font.Size = 13;
                            cell.Range.ParagraphFormat.SpaceAfter = 0;
                            cell.Range.ParagraphFormat.SpaceBefore = 0;
                            cell.TopPadding = 0;
                            cell.BottomPadding = 0;
                            cell.LeftPadding = 3;
                            cell.RightPadding = 3;
                            cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                            if (r == 1 || (tableConfig.name == "Результаты испытаний" && r == 2))
                            {
                                cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            }
                            else
                            {
                                if (c == 1)
                                    cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                                else
                                    cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                            }
                        }
                    }

                    // Высота строк
                    foreach (Word.Row row in newTable.Rows)
                    {
                        float minHeight = InchesToPoints(0.2f);
                        if (minHeight >= 1 && minHeight <= 1000)
                        {
                            row.HeightRule = Word.WdRowHeightRule.wdRowHeightAtLeast;
                            row.Height = minHeight;
                        }
                        else
                        {
                            row.HeightRule = Word.WdRowHeightRule.wdRowHeightAuto;
                        }
                    }

                    // Отступ после таблицы
                    Word.Range afterTable = newTable.Range;
                    afterTable.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    afterTable.InsertAfter("\n");

                    MessageBox.Show($"✅ Таблица '{tableConfig.name}' успешно создана по закладке '{tableConfig.bookmark}'.", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"❌ Ошибка создания таблицы '{tableConfig.name}': {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private float InchesToPoints(float inches)
        {
            return Math.Max(1, inches * 72);
        }
    }
}