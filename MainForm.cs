using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace TitleGen
{
    public class MainForm : Form
    {
        private Panel testsPanel;
        private Panel inputsPanel;
        private RadioButton radioTip;
        private RadioButton radioPeriod;
        private RadioButton radioTest;
        private TextBox txtTemplate;
        private Button btnGenerate;

        private Dictionary<string, TextBox> inputs = new Dictionary<string, TextBox>();

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
            // Панель слева с испытаниями
            testsPanel = new Panel()
            {
                Left = 10,
                Top = 10,
                Width = 250,
                Height = 500,
                BorderStyle = BorderStyle.FixedSingle,
                AutoScroll = true
            };
            this.Controls.Add(testsPanel);

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
                y += 25;
            }

            // Радиокнопки выбора шаблона
            radioTip = new RadioButton()
            {
                Text = "Типовые",
                Left = 280,
                Top = 20,
                AutoSize = true
            };
            radioTip.CheckedChanged += TemplateSelectorChanged;
            this.Controls.Add(radioTip);

            radioPeriod = new RadioButton()
            {
                Text = "Периодические",
                Left = 380,
                Top = 20,
                AutoSize = true
            };
            radioPeriod.CheckedChanged += TemplateSelectorChanged;
            this.Controls.Add(radioPeriod);

            radioTest = new RadioButton()
            {
                Text = "Тест",
                Left = 520,
                Top = 20,
                AutoSize = true
            };
            radioTest.CheckedChanged += TemplateSelectorChanged;
            this.Controls.Add(radioTest);

            // Поле для пути к шаблону
            txtTemplate = new TextBox()
            {
                Left = 280,
                Top = 60,
                Width = 500
            };
            this.Controls.Add(txtTemplate);

            // Панель для динамических inputbox
            inputsPanel = new Panel()
            {
                Left = 280,
                Top = 100,
                Width = 500,
                Height = 350,
                AutoScroll = true,
                BorderStyle = BorderStyle.FixedSingle
            };
            this.Controls.Add(inputsPanel);

            // Кнопка генерации DOCX
            btnGenerate = new Button()
            {
                Text = "Сформировать DOCX",
                Left = 280,
                Top = 470,
                Width = 200
            };
            btnGenerate.Click += btnGenerate_Click;
            this.Controls.Add(btnGenerate);
        }

        /// <summary>
        /// Смена шаблона при выборе радиокнопки
        /// </summary>
        private void TemplateSelectorChanged(object sender, EventArgs e)
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;

            if (radioTip.Checked)
                txtTemplate.Text = Path.Combine(baseDir, "tipovye.docx");
            else if (radioPeriod.Checked)
                txtTemplate.Text = Path.Combine(baseDir, "periodich.docx");
            else if (radioTest.Checked)
                txtTemplate.Text = Path.Combine(baseDir, "test.docx");

            if (File.Exists(txtTemplate.Text))
                BuildDynamicForm(txtTemplate.Text);
            else
                inputsPanel.Controls.Clear();
        }

        /// <summary>
        /// Чтение плейсхолдеров из шаблона и создание inputbox
        /// </summary>
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

        /// <summary>
        /// Извлечение плейсхолдеров {{name}} из Word-документа
        /// </summary>
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
                if (doc != null) doc.Close(false);
                wordApp.Quit(false);
            }

            return placeholders;
        }

        /// <summary>
        /// Генерация итогового DOCX
        /// </summary>
        private void btnGenerate_Click(object sender, EventArgs e)
        {
            if (!File.Exists(txtTemplate.Text))
            {
                MessageBox.Show("Шаблон не найден!");
                return;
            }

            string output = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "Протокол.docx"
            );

            Word.Application wordApp = new Word.Application();
            Word.Document doc = null;
            try
            {
                doc = wordApp.Documents.Open(txtTemplate.Text);

                foreach (var pair in inputs)
                {
                    string placeholder = "{{" + pair.Key + "}}";
                    string value = pair.Value.Text;

                    Word.Find findObject = wordApp.Selection.Find;
                    findObject.ClearFormatting();
                    findObject.Text = placeholder;
                    findObject.Replacement.ClearFormatting();
                    findObject.Replacement.Text = value;

                    object replaceAll = Word.WdReplace.wdReplaceAll;
                    findObject.Execute(Replace: ref replaceAll);
                }

                doc.SaveAs2(output);
                MessageBox.Show("DOCX создан: " + output);
            }
            finally
            {
                if (doc != null) doc.Close();
                wordApp.Quit();
            }
        }
    }
}
