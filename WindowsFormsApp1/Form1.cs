using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.IO;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public double result; // Переменная для хранения результата
        public string additionally = "не выбрано"; // Строка для хранения дополнительной информации

        public Form1()
        {
            InitializeComponent();
        }

        // Метод для проверки номера
        public void NumberCheck()
        {
            int numberCheck;
            Random random = new Random();
            numberCheck = random.Next(1, 9);
        }

        // Обработчик нажатия на кнопку "Рассчитать"
        private void button1_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear(); // Очистка списка результатов
            try
            {
                double height = Convert.ToDouble(textBox1.Text);
                double width = Convert.ToDouble(textBox2.Text);

                if (textBox1.Text == "" || textBox2.Text == "")
                {
                    MessageBox.Show("Поля должны быть заполнены");
                    return;
                }

                if (radioButton1.Checked)
                {
                    result = height * width * 213.15;
                }
                else if (radioButton2.Checked)
                {
                    result = width * height * 265.80;
                }

                // Обработка выбранных чекбоксов
                if (checkBox1.Checked && checkBox2.Checked)
                {
                    result *= 1.3 * 1.26;
                    additionally = "Многоуровневый и Фотопечать";
                }
                else if (checkBox1.Checked)
                {
                    result *= 1.3;
                    additionally = "Многоуровневый";
                }
                else if (checkBox2.Checked)
                {
                    result *= 1.26;
                    additionally = "Фотопечать";
                }
                else
                {
                    additionally = "не выбрано";
                }

                result = Math.Round(result, 2);
                listBox1.Items.Add(result);
            }
            catch (FormatException ex)
            {
                MessageBox.Show("Ошибка");
            }
        }

        // Обработчик нажатия на кнопку "Создать квитанцию"
        private void button2_Click(object sender, EventArgs e)
        {
            CreateReceipt(); // Вызов метода для создания квитанции
        }

        // Метод для создания квитанции
        public void CreateReceipt()
        {
            string date = DateTime.Now.ToString(); // Получение текущей даты и времени
            int number;
            Random rnd = new Random();
            number = rnd.Next(10000, 100000); // Генерация случайного номера

            // Путь к вашему шаблону квитанции
            string templatePath = @"C:\Users\User\Desktop\Квитанции\Квтанция_шаблон.docx";

            if (!File.Exists(templatePath))
            {
                MessageBox.Show("Файл шаблона не найден");
                return;
            }

            // Создание экземпляра Word и открытие документа
            Word.Application wordApp = new Word.Application();
            Word.Document doc = wordApp.Documents.Open(templatePath);

            // Создаем новый файл с уникальным именем
            string newFileName = $"Квитанция{DateTime.Now.ToString("yyyyMMddHHmmss")}.docx";
            string newFilePath = Path.Combine(Path.GetDirectoryName(templatePath), newFileName);

            // Замена текста в документе
            ReplaceText(doc, "[number]", number.ToString());
            ReplaceText(doc, "[height]", textBox1.Text.ToString());
            ReplaceText(doc, "[weight]", textBox2.Text.ToString());
            ReplaceText(doc, "[result]", result.ToString());
            ReplaceText(doc, "[date]", date);
            ReplaceText(doc, "[additionally]", additionally);

            // Сохранение как нового файла
            doc.SaveAs(newFilePath);
            wordApp.Visible = true;
        }

        // Метод для замены текста в документе Word
        private void ReplaceText(Word.Document doc, string findText, string replaceText)
        {
            Word.Range range = doc.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: findText, ReplaceWith: replaceText);
        }
    }
}
