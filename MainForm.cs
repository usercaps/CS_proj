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

        // Параметры (старый UI)
        private Panel testsPanel;
        private Panel inputsPanel;
        private RadioButton radioTip;
        private RadioButton radioPeriod;
        private RadioButton radioTest;
        private TextBox txtTemplate;
        private Button btnGenerate;

        private Dictionary<string, TextBox> inputs = new Dictionary<string, TextBox>();
        private Dictionary<string, CheckBox> testCheckboxes = new Dictionary<string, CheckBox>();

        // Редактор таблиц
        private ComboBox cmbTables;
        private DataGridView dgvRows;
        private Button btnAddRow, btnDeleteRow, btnSaveConfig;
        private TemplateConfig currentConfig;
        private string currentConfigPath;
        private TableConfig currentTable;

        // 👇 Список глобально общих СИ/ИО — добавляются один раз
        private List<TableRow> commonEquipment = new List<TableRow>
        {
            new TableRow
            {
                testName = "*",
                values = new List<string> { "", "Барометр БАММ-1", "Б-001", "2025-12-31" }
            },
            new TableRow
            {
                testName = "*",
                values = new List<string> { "", "Термометр ВИТ-1", "ВИТ-001", "2025-11-15" }
            },
            new TableRow
            {
                testName = "*",
                values = new List<string> { "", "Гигрометр ВИТ-2", "Г-002", "2025-10-20" }
            }
        };

        // 👇 Группы испытаний и их общие приборы
        private Dictionary<string, List<string>> testGroups = new Dictionary<string, List<string>>
        {
            { "Температура", new List<string> { "Повышенная температура", "Пониженная температура", "Циклы температуры" } },
            { "Давление", new List<string> { "Давление рабочее", "Давление предельное" } },
            { "Влажность", new List<string> { "Повышенная влажность", "Пониженная влажность" } }
        };

        private Dictionary<string, TableRow> groupEquipment = new Dictionary<string, TableRow>
        {
            { "Температура", new TableRow
                {
                    testName = "Температура",
                    values = new List<string> { "", "Термокамера Binder", "TK-2024-001", "2025-12-01" }
                }
            },
            { "Давление", new TableRow
                {
                    testName = "Давление",
                    values = new List<string> { "", "Манометр МД-100", "МД-001", "2025-11-30" }
                }
            },
            { "Влажность", new TableRow
                {
                    testName = "Влажность",
                    values = new List<string> { "", "Камера влажности Climats", "CV-2024", "2025-10-15" }
                }
            }
        };

        public MainForm()
        {
            this.Text = "Генерация протокола (DocX)";
            this.Width = 850;
            this.Height = 600;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.AutoScroll = true;

            BuildStaticUI();
        }

        private void BuildStaticUI()
        {
            tabControl = new TabControl()
            {
                Left = 10,
                Top = 10,
                Width = 820,
                Height = 550
            };

            tabParams = new TabPage() { Text = "Параметры" };
            BuildParamsTab(tabParams);
            tabControl.TabPages.Add(tabParams);

            tabTableEditor = new TabPage() { Text = "Редактор таблиц" };
            BuildTableEditorTab(tabTableEditor);
            tabControl.TabPages.Add(tabTableEditor);

            this.Controls.Add(tabControl);
        }

        private void BuildParamsTab(TabPage page)
        {
            testsPanel = new Panel()
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
                CheckBox cb = new CheckBox()
                {
                    Text = test,
                    Left = 10,
                    Top = y,
                    AutoSize = true
                };
                testsPanel.Controls.Add(cb);
                testCheckboxes[test] = cb;
                y += 25;
            }

            // 👇 Подписываемся на изменение чекбоксов
            foreach (var cb in testCheckboxes.Values)
            {
                cb.CheckedChanged += (s, ev) => UpdateRowStatuses();
            }

            radioTip = new RadioButton()
            {
                Text = "Типовые",
                Left = 280,
                Top = 20,
                AutoSize = true
            };
            radioTip.CheckedChanged += TemplateSelectorChanged;
            page.Controls.Add(radioTip);

            radioPeriod = new RadioButton()
            {
                Text = "Периодические",
                Left = 380,
                Top = 20,
                AutoSize = true
            };
            radioPeriod.CheckedChanged += TemplateSelectorChanged;
            page.Controls.Add(radioPeriod);

            radioTest = new RadioButton()
            {
                Text = "Тест",
                Left = 520,
                Top = 20,
                AutoSize = true
            };
            radioTest.CheckedChanged += TemplateSelectorChanged;
            page.Controls.Add(radioTest);

            txtTemplate = new TextBox()
            {
                Left = 280,
                Top = 60,
                Width = 500
            };
            page.Controls.Add(txtTemplate);

            inputsPanel = new Panel()
            {
                Left = 280,
                Top = 100,
                Width = 500,
                Height = 300,
                AutoScroll = true,
                BorderStyle = BorderStyle.FixedSingle
            };
            page.Controls.Add(inputsPanel);

            btnGenerate = new Button()
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
            Label lblTable = new Label()
            {
                Text = "Выберите таблицу:",
                Left = 20,
                Top = 20,
                AutoSize = true
            };
            page.Controls.Add(lblTable);

            cmbTables = new ComboBox()
            {
                Left = 150,
                Top = 18,
                Width = 300,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbTables.SelectedIndexChanged += cmbTables_SelectedIndexChanged;
            page.Controls.Add(cmbTables);

            dgvRows = new DataGridView()
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

            btnAddRow = new Button()
            {
                Text = "Добавить строку",
                Left = 20,
                Top = 420,
                Width = 150
            };
            btnAddRow.Click += btnAddRow_Click;
            page.Controls.Add(btnAddRow);

            btnDeleteRow = new Button()
            {
                Text = "Удалить строку",
                Left = 180,
                Top = 420,
                Width = 150
            };
            btnDeleteRow.Click += btnDeleteRow_Click;
            page.Controls.Add(btnDeleteRow);

            btnSaveConfig = new Button()
            {
                Text = "Сохранить config.json",
                Left = 600,
                Top = 420,
                Width = 180
            };
            btnSaveConfig.Click += btnSaveConfig_Click;
            page.Controls.Add(btnSaveConfig);
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

                try
                {
                    string json = JsonConvert.SerializeObject(currentConfig, Formatting.Indented);
                    File.WriteAllText(configPath, json);
                    MessageBox.Show($"Создан новый config.json по умолчанию:\n{configPath}", "Инфо", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Не удалось создать config.json:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                try
                {
                    string json = File.ReadAllText(configPath);
                    currentConfig = JsonConvert.DeserializeObject<TemplateConfig>(json) ?? new TemplateConfig();
                    currentConfigPath = configPath;
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка загрузки config.json:\n{ex.Message}\n\nСоздан конфиг по умолчанию.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                        columns = new List<string> { "№", "Наименование испытания", "Методика", "Требования" },
                        rows = new List<TableRow>
                        {
                            new TableRow
                            {
                                testName = "Повышенная температура",
                                values = new List<string> { "1", "Повышенная температура", "ГОСТ 12345-67", "Выдержать 72ч при +85°C" }
                            },
                            new TableRow
                            {
                                testName = "Вибрация",
                                values = new List<string> { "2", "Вибрация", "ГОСТ 30630.2.1", "Частота 10-55 Гц, амплитуда 1.5 мм" }
                            },
                            new TableRow
                            {
                                testName = "Безопасность",
                                values = new List<string> { "3", "Безопасность", "ГОСТ Р МЭК 61010", "Нет пробоя изоляции" }
                            }
                        }
                    },
                    new TableConfig
                    {
                        name = "СИ и ИО",
                        bookmark = "Table_Equipment",
                        columns = new List<string> { "№", "Наименование СИ/ИО", "Зав. №", "Поверка до" },
                        rows = new List<TableRow>
                        {
                            new TableRow
                            {
                                testName = "Вибрация",
                                values = new List<string> { "", "Вибростенд LDS V408", "VS-408-001", "2025-11-30" }
                            }
                        }
                    },
                    new TableConfig
                    {
                        name = "Результаты испытаний",
                        bookmark = "Table_Results",
                        columns = new List<string> { "№", "Наименование испытания", "Результат", "Примечание" },
                        rows = new List<TableRow>
                        {
                            new TableRow
                            {
                                testName = "Повышенная температура",
                                values = new List<string> { "1", "Повышенная температура", "", "" }
                            },
                            new TableRow
                            {
                                testName = "Вибрация",
                                values = new List<string> { "2", "Вибрация", "", "" }
                            },
                            new TableRow
                            {
                                testName = "Безопасность",
                                values = new List<string> { "3", "Безопасность", "", "" }
                            }
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
            {
                cmbTables.Items.Add(table.name);
            }

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
            dgvRows.Columns.Add("status", "Статус"); // 👈 новая колонка

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
                {
                    status = "Скрыто";
                }
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
            if (dgvRows.Columns["testName"] is DataGridViewComboBoxColumn)
                return;

            var comboBoxColumn = new DataGridViewComboBoxColumn
            {
                Name = "testName",
                HeaderText = "Привязка к чекбоксу"
            };

            foreach (var testName in testCheckboxes.Keys)
            {
                comboBoxColumn.Items.Add(testName);
            }

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
                {
                    status = "Скрыто";
                }

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
            if (currentTable == null)
            {
                MessageBox.Show("Выберите таблицу для редактирования.");
                return;
            }

            dgvRows.Rows.Add();

            if (testCheckboxes.Count > 0)
            {
                dgvRows.Rows[dgvRows.Rows.Count - 1].Cells["testName"].Value = testCheckboxes.Keys.First();
            }
        }

        private void btnDeleteRow_Click(object sender, EventArgs e)
        {
            if (dgvRows.SelectedRows.Count == 0)
            {
                MessageBox.Show("Выберите строку для удаления.");
                return;
            }

            dgvRows.Rows.RemoveAt(dgvRows.SelectedRows[0].Index);
        }

        private void btnSaveConfig_Click(object sender, EventArgs e)
        {
            if (currentConfig == null || currentTable == null)
            {
                MessageBox.Show("Нет данных для сохранения.");
                return;
            }

            currentTable.rows = new List<TableRow>();
            foreach (DataGridViewRow row in dgvRows.Rows)
            {
                if (row.IsNewRow) continue;

                var values = new List<string>();
                for (int i = 1; i < row.Cells.Count; i++)
                {
                    if (i == 1) continue; // Пропускаем колонку "Статус"
                    values.Add(row.Cells[i].Value?.ToString() ?? "");
                }

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
                Label lbl = new Label()
                {
                    Text = ph,
                    Left = 10,
                    Top = y + 3,
                    Width = 200
                };
                inputsPanel.Controls.Add(lbl);

                TextBox tb = new TextBox()
                {
                    Left = 220,
                    Top = y,
                    Width = 250
                };
                inputsPanel.Controls.Add(tb);

                inputs[ph] = tb;
                y += 30;
            }
        }

        private List<string> ExtractPlaceholders(string path)
        {
            List<string> placeholders = new List<string>();

            Word.Application wordApp = new Word.Application();
            Word.Document doc = null;
            try
            {
                doc = wordApp.Documents.Open(path, ReadOnly: true, Visible: false);
                string text = doc.Content.Text;

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
                if (doc != null)
                {
                    doc.Close(false);
                    Marshal.ReleaseComObject(doc);
                }

                if (wordApp != null)
                {
                    wordApp.Quit(false);
                    Marshal.ReleaseComObject(wordApp);
                }

                doc = null;
                wordApp = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();
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
                MessageBox.Show("Конфиг config.json не найден рядом с шаблоном!");
                return;
            }

            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "Word Document (*.docx)|*.docx";
                sfd.Title = "Сохранить протокол как...";
                sfd.FileName = "Протокол.docx";

                if (sfd.ShowDialog() != DialogResult.OK)
                    return;

                string output = sfd.FileName;

                Word.Application wordApp = null;
                Word.Document doc = null;
                try
                {
                    wordApp = new Word.Application();
                    doc = wordApp.Documents.Open(txtTemplate.Text, ReadOnly: false, Visible: false);

                    ReplacePlaceholdersInDocument(doc);

                    string json = File.ReadAllText(configPath);
                    var config = JsonConvert.DeserializeObject<TemplateConfig>(json);

                    ProcessTablesFromConfig(doc, config);

                    ReplacePlaceholdersInDocument(doc);

                    doc.SaveAs2(output);
                    MessageBox.Show("DOCX успешно создан:\n" + output);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка при генерации:\n" + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    if (doc != null)
                    {
                        doc.Close(false);
                        Marshal.ReleaseComObject(doc);
                    }

                    if (wordApp != null)
                    {
                        wordApp.Quit();
                        Marshal.ReleaseComObject(wordApp);
                    }

                    doc = null;
                    wordApp = null;

                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
        }

        private void ReplacePlaceholdersInDocument(Word.Document doc)
        {
            foreach (var pair in inputs)
            {
                string placeholder = "{{" + pair.Key + "}}";
                string value = pair.Value.Text;

                Word.Range range = doc.Content;
                range.Find.ClearFormatting();
                range.Find.Execute(
                    FindText: placeholder,
                    ReplaceWith: value,
                    Replace: Word.WdReplace.wdReplaceAll
                );
            }
        }

        private void ProcessTablesFromConfig(Word.Document doc, TemplateConfig config)
        {
            foreach (var tableConfig in config.tables)
            {
                if (!doc.Bookmarks.Exists(tableConfig.bookmark))
                    continue;

                Word.Bookmark bm = doc.Bookmarks[tableConfig.bookmark];
                if (!(bm.Range.Tables.Count > 0))
                    continue;

                Word.Table table = bm.Range.Tables[1];

                if (table.Rows.Count > 1)
                {
                    try
                    {
                        table.Rows[2].Delete();
                    }
                    catch { /* игнорируем */ }
                }

                var rowsToInsert = new List<TableRow>();

                // 1. Индивидуальные строки из JSON
                foreach (var row in tableConfig.rows)
                {
                    if (testCheckboxes.TryGetValue(row.testName, out CheckBox cb) && cb.Checked)
                    {
                        rowsToInsert.Add(row);
                    }
                }

                // 2. Общие СИ/ИО — добавляем ОДИН РАЗ, без дублей
                if (tableConfig.bookmark == "Table_Equipment")
                {
                    // Глобально общие приборы
                    string anyTest = testCheckboxes
                        .Where(kvp => kvp.Value.Checked)
                        .Select(kvp => kvp.Key)
                        .FirstOrDefault() ?? "";

                    foreach (var commonRow in commonEquipment)
                    {
                        var clonedRow = new TableRow
                        {
                            testName = anyTest,
                            values = new List<string>(commonRow.values)
                        };
                        rowsToInsert.Add(clonedRow);
                    }

                    // Групповые приборы
                    foreach (var group in testGroups)
                    {
                        var groupName = group.Key;
                        var groupTests = group.Value;

                        var selectedTestsInGroup = groupTests
                            .Where(test => testCheckboxes.ContainsKey(test) && testCheckboxes[test].Checked)
                            .ToList();

                        if (selectedTestsInGroup.Any() && groupEquipment.ContainsKey(groupName))
                        {
                            var groupRow = groupEquipment[groupName];
                            var targetTest = selectedTestsInGroup.First();

                            var clonedRow = new TableRow
                            {
                                testName = targetTest,
                                values = new List<string>(groupRow.values)
                            };
                            rowsToInsert.Add(clonedRow);
                        }
                    }
                }

                // Вставляем строки
                for (int i = rowsToInsert.Count - 1; i >= 0; i--)
                {
                    var rowData = rowsToInsert[i];

                    Word.Row newRow;
                    if (table.Rows.Count > 1)
                    {
                        Word.Row lastRow = table.Rows.Last;
                        Word.Range range = lastRow.Range;
                        newRow = table.Rows.Add(range);
                    }
                    else
                    {
                        newRow = table.Rows.Add();
                    }

                    for (int colIndex = 0; colIndex < rowData.values.Count && colIndex < newRow.Cells.Count; colIndex++)
                    {
                        newRow.Cells[colIndex + 1].Range.Text = rowData.values[colIndex];
                    }
                }

                // Автонумерация
                if (tableConfig.bookmark == "Table_Equipment")
                {
                    int rowNum = 1;
                    foreach (Word.Row row in table.Rows)
                    {
                        if (row.Index == 1) continue;
                        if (row.Cells.Count > 0)
                        {
                            row.Cells[1].Range.Text = rowNum.ToString();
                            rowNum++;
                        }
                    }
                }
            }
        }
    }
}