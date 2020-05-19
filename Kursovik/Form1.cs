using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Runtime.CompilerServices;

namespace Kursovik
{
    public partial class Form1 : Form
    {
        DataTable dt = new DataTable();
        public Form1()
        {
            InitializeComponent();
        }


        private void button1_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                //парсим x и y
                double x = Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value);
                double y = Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value);

                // x^2
                dataGridView1.Rows[i].Cells[2].Value = Math.Round(Math.Pow(x, 2), 2);

                // y^2
                dataGridView1.Rows[i].Cells[3].Value = Math.Round(Math.Pow(y, 2), 2);

                // x*y
                dataGridView1.Rows[i].Cells[4].Value = Math.Round(x * y, 2);
            }

            //количество входных данных
            int countFirst = dataGridView1.Rows.Count;

            //сумма столбцов 
            double sumX = 0,
                    sumY = 0,
                    sumX2 = 0,
                    sumY2 = 0,
                    sumXY = 0;

            for (int i = 0; i < countFirst; i++)
            {
                sumX += Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value);
                sumY += Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value);
                sumX2 += Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value);
                sumY2 += Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value);
                sumXY += Convert.ToDouble(dataGridView1.Rows[i].Cells[4].Value);
            }

            dt.Rows.Add(sumX, sumY, sumX2, sumY2, sumXY);
            dataGridView1.Rows[dataGridView1.Rows.Count - 1].HeaderCell.Value = "Σ";

            dt.Rows.Add(Math.Round(sumX / countFirst, 2), Math.Round(sumY / countFirst, 2), Math.Round(sumX2 / countFirst, 2),
                Math.Round(sumY2 / countFirst, 2), Math.Round(sumXY / countFirst, 2));
            dataGridView1.Rows[dataGridView1.Rows.Count - 1].HeaderCell.Value = "Средняя \r\nвеличина";


        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dt.Columns.Add("x");
            dt.Columns.Add("y");
            dt.Columns.Add("x2");
            dt.Columns.Add("y2");
            dt.Columns.Add("xy");
            dataGridView1.DataSource = dt;
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog file = new OpenFileDialog();
            if (file.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = file.FileName;
                //получаем количество строк в файле
                string[] lines = File.ReadAllLines(file.FileName);
                //массив для значений
                string[] values;

                for (int i = 0; i < lines.Length; i++)
                {
                    values = lines[i].ToString().Split(';');
                    string[] row = new string[values.Length];

                    for (int j = 0; j < values.Length; j++)
                    {
                        row[j] = values[j].Trim();
                    }

                    dt.Rows.Add(row);

                }
            }
        }
    }
}
