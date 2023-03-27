using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace Сиргиенко_Софья_экзаменн_ПМ_02
{
    public partial class Form1 : Form
    {
        string path = "";
        string NameFile;

        public Form1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try 
            {
                if (radioButton1.Checked && !radioButton2.Checked || checkBox1.Checked && !checkBox2.Checked)
                {
                    double rez = Convert.ToDouble(textBox1.Text) * Convert.ToDouble(textBox2.Text) * 213.15;
                    double discount = rez / 100 * 26;
                    double itog = rez + discount;
                    label3.Text = "Стоимость заказа " + itog.ToString("F2") + " рублей";
                }
                else if (!radioButton1.Checked && radioButton2.Checked || checkBox1.Checked && !checkBox2.Checked)
                {
                    double rez = Convert.ToDouble(textBox1.Text) * Convert.ToDouble(textBox2.Text) * 265.80;
                    double discount = rez / 100 * 26;
                    double itog = rez + discount;
                    label3.Text = "Стоимость заказа " + itog.ToString("F2") + " рублей";
                }
                else if (radioButton1.Checked && !radioButton2.Checked || !checkBox1.Checked && checkBox2.Checked)
                {
                    double rez = Convert.ToDouble(textBox1.Text) * Convert.ToDouble(textBox2.Text) * 213.15;
                    double discount = rez / 100 * 30;
                    double itog = rez + discount;
                    label3.Text = "Стоимость заказа " + itog.ToString("F2") + " рублей";
                }
                else if (!radioButton1.Checked && radioButton2.Checked || checkBox1.Checked && !checkBox2.Checked)
                {
                    double rez = Convert.ToDouble(textBox1.Text) * Convert.ToDouble(textBox2.Text) * 265.80;
                    double discount = rez / 100 * 30;
                    double itog = rez + discount;
                    label3.Text = "Стоимость заказа " + itog.ToString("F2") + " рублей";
                }
                else if (!radioButton1.Checked && !radioButton2.Checked || !checkBox1.Checked && !checkBox2.Checked)
                {
                    MessageBox.Show("Введите все данные!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                else
                    MessageBox.Show("Введите все данные!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //загрузка картинки
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Filter = "Text files(*.png)|*.txt|All files(*.jpg)|*.*";
            openFile.ShowDialog();

            path = openFile.FileName;

            pictureBox1.ImageLocation = path;
            NameFile = Path.GetFileName(path);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            // Создаём объект документа
            Document doc = null;
            try
            {
                // Создаём объект приложения
                Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
                // Путь до шаблона документа
                string source = Environment.CurrentDirectory + @"\чек.docx";
                // Открываем
                doc = app.Documents.Open(source);
                doc.Activate();

                // Добавляем информацию
                // wBookmarks содержит все закладки
                Bookmarks wBookmarks = doc.Bookmarks;
                Range wRange;
                int i = 0;
                string[] data = new string[2] {checkBox1.Text, label3.Text};
                foreach (Bookmark mark in wBookmarks)
                {
                    wRange = mark.Range;
                    wRange.Text = data[i];
                    i++;
                }

                // Закрываем документ
                doc.Close();
                doc = null;
            }
            catch (Exception ex)
            {
                doc.Close();
                doc = null;
                MessageBox.Show(ex.Message);
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;

            if (!Char.IsDigit(number))
            {
                e.Handled = true;
            }
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;

            if (!Char.IsDigit(number))
            {
                e.Handled = true;
            }
        }
    }
}

