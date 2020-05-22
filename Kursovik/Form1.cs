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
using System.Windows.Forms.DataVisualization.Charting;

namespace Kursovik
{
    public partial class Form1 : Form
    {
        DataTable dt = new DataTable();
        Boolean check_click = false;
        public Form1()
        {
            InitializeComponent();
            
        }

        //метод для обработки неправильно введенных данных
        private bool checkError()
        {

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                //парсим x и y
                try
                {
                    double x = Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value);
                    double y = Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value);
                }
                catch (FormatException err)
                {
                    MessageBox.Show("Неверный формат входных данных!\n\rНа строке №" + Convert.ToString(i + 1));
                    //фокус курсора на строке с неверными данными
                    dataGridView1.ClearSelection();
                    dataGridView1.Rows[i].Selected = true;
                    dataGridView1.CurrentCell = dataGridView1[0, i];
                    return false;
                }
            }
            return true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            if (checkError())
            {               
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {

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

                //добавление строки с суммами
                dt.Rows.Add(sumX, sumY, sumX2, sumY2, sumXY);
                dataGridView1.Rows[dataGridView1.Rows.Count - 1].HeaderCell.Value = "Σ";

                //добавление строки с средними величинами
                double avgX = Math.Round(sumX / countFirst, 2),
                       avgY = Math.Round(sumY / countFirst, 2),
                       avgX2 = Math.Round(sumX2 / countFirst, 2),
                       avgY2 = Math.Round(sumY2 / countFirst, 2),
                       avgXY = Math.Round(sumXY / countFirst, 2);
                dt.Rows.Add(avgX, avgY, avgX2, avgY2, avgXY);
                dataGridView1.Rows[dataGridView1.Rows.Count - 1].HeaderCell.Value = "Средняя \r\nвеличина";

                //добавление строки с средними квадратическими отклонениями для признаков x и y соответственно
                double kvX = Math.Round(Math.Sqrt(sumX2 / countFirst - Math.Pow(sumX / countFirst, 2)), 2);
                double kvY = Math.Round(Math.Sqrt(sumY2 / countFirst - Math.Pow(sumY / countFirst, 2)), 2);
                dt.Rows.Add(kvX, kvY);
                dataGridView1.Rows[dataGridView1.Rows.Count - 1].HeaderCell.Value = "Средние \r\nотклонения";

                //линейный коэффициент корреляции
                double koef_korr = (avgXY - (sumX / countFirst) * (sumY / countFirst)) / (kvX * kvY);
                string koefMsg = "",
                       koefZnak = "";

                if (koef_korr == 0)
                {
                    koefMsg = "Связь отсутствует";
                }
                if ((Math.Abs(koef_korr) > 0) && (Math.Abs(koef_korr) < 0.3))
                {
                    koefMsg = "Слабая";
                }
                if ((Math.Abs(koef_korr) >= 0.3) && (Math.Abs(koef_korr) <= 0.7))
                {
                    koefMsg = "Средней силы";
                }
                if ((Math.Abs(koef_korr) > 0.7) && (Math.Abs(koef_korr) < 1))
                {
                    koefMsg = "Сильная";
                }
                if (koef_korr == 1)
                {
                    koefMsg = "Функциональная";
                }
                if ((koef_korr > 0) && (koef_korr < 1))
                {
                    koefZnak = "Прямая";
                }
                if ((koef_korr < 0) && (koef_korr > -1))
                {
                    koefZnak += "Обратная";
                }
                dt.Rows.Add(Math.Round(koef_korr, 2), koefMsg, koefZnak);
                dataGridView1.Rows[dataGridView1.Rows.Count - 1].HeaderCell.Value = "Коэффициент \r\nкорреляции,\r\nr";

                //ошибка коэфф корреляции m
                double errKoefKorr = Math.Round(Math.Abs(koef_korr/Math.Sqrt((1 - Math.Round(koef_korr * koef_korr, 2)) / (countFirst - 2))), 2);
                dt.Rows.Add(errKoefKorr);
                dataGridView1.Rows[dataGridView1.Rows.Count - 1].HeaderCell.Value = "Критерий \r\nдоверенности,\r\nt";

                //метод для получения Критерия Стюдента
                //____________________________________ какой уровень значимости выбрать не знаю (выбрал как в лекциях - 0,05)
                Chart chart1 = new Chart();
                double koef_znach = 0.05;
                double t_tabl = Math.Round(chart1.DataManipulator.Statistics.InverseTDistribution(koef_znach, countFirst - 2), 2);
                dt.Rows.Add(t_tabl,koef_znach);
                dataGridView1.Rows[dataGridView1.Rows.Count - 1].HeaderCell.Value = "Критерий \r\nСтьюдента,\r\nпри коэффициенте значимости";

                //делаем вывод после сравнения ошибки с Стюдента
                string vivod_korr = "Коэфф ошибки корреляции и Стюдента совпадают";

                if (errKoefKorr>t_tabl)
                {
                    vivod_korr = "Связь присутствует";
                }
                if (errKoefKorr<t_tabl)
                {
                    vivod_korr = "Связь случайна";
                }
                dt.Rows.Add(vivod_korr);
                dataGridView1.Rows[dataGridView1.Rows.Count - 1].HeaderCell.Value = "Вывод \r\nо значимости коэффициента корреляции";



                //рисование графика
                double[] X = new double[countFirst];
                double[] Y = new double[countFirst];

                //поиск коэфф регрессии
                double k = Math.Round((countFirst * sumXY - sumX * sumY) / (countFirst * sumX2 - sumX * sumX), 2);
                double b = Math.Round(avgY - k * avgX, 2);


                for (int i = 0; i < countFirst; i++)
                {
                    //парсим значения Х и У
                    X[i] = Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value);
                    Y[i] = Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value);
                  
                    //добавляем точки и линию на график
                    chart2.Series["Series1"].Points.AddXY(X[i],Y[i]);
                    chart2.Series["line"].Points.AddXY(X[i], k*X[i]+b);          
                }
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            dt.Columns.Add("x");
            dt.Columns.Add("y");
            dt.Columns.Add("x2");
            dt.Columns.Add("y2");
            dt.Columns.Add("xy");
            dataGridView1.DataSource = dt;

            //выбираем тип диаграммы точечный
            chart2.Series["Series1"].ChartType = SeriesChartType.Point;

            //выбираем тип диаграммы линейный
            chart2.Series.Add("line");
            chart2.Series["line"].ChartType = SeriesChartType.Line;

        }       

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //очистка таблицы
            dt.Clear();
            
            //очистка графика
            foreach (var series in chart2.Series)
            {
                series.Points.Clear();
            }

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
                    dataGridView1.Rows[dataGridView1.Rows.Count-1].HeaderCell.Value = Convert.ToString(i+1);
                }
                checkError();
            }      
        }
    }
}
