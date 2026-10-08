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
using Xceed.Words.NET;

namespace TextEditor
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void новыйToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();
        }

        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Текстовые файлы (*.txt)|*.txt|" +
                "Все файлы (*.*)|*.*";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    string filename = openFileDialog1.FileName;
                    string text = File.ReadAllText(filename);
                    richTextBox1.Text = text;
                } 
                catch
                {
                    MessageBox.Show("Не удалось открыть файл!", "Ошибка",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void сохранитьКакToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "Текстовые файлы (*.txt)|*.txt|" +
                "Word документы (*.docx)|*.docx";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string filename = saveFileDialog1.FileName;
                try
                {
                    if (saveFileDialog1.FilterIndex == 1)
                    {
                        File.WriteAllText(filename, richTextBox1.Text);
                    }
                    else if (saveFileDialog1.FilterIndex == 2)
                    {
                        using (DocX document = DocX.Create(filename))
                        {
                            document.InsertParagraph(richTextBox1.Text);
                            document.Save();
                        }
                    }

                    MessageBox.Show("Файл успешно сохранен!", "Сохранение", MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                       "Не удалось сохранить файл!\n\n" +
                       ex.Message,
                       "Ошибка",
                       MessageBoxButtons.OK,
                       MessageBoxIcon.Error);
                }
            }
        }

        private void выделитьВсеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectAll();
        }

        private void копироватьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Copy();
        }

        private void вырезатьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Cut();
        }

        private void вставитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Paste();
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void справкаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Простой текстовый редактор\n" +
                "Windows Forms\n", "О программе");
        }

        private void открытьФормуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form2 form2 = new Form2();
            Form3 form3 = new Form3();
            form2.Show();
            form3.Show();
            //form2.ShowDialog();
        }
    }
}
