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
using System.Diagnostics;
using Word = Microsoft.Office.Interop.Word;

namespace Kursovik
{
    public partial class Form1 : Form
    {
        DataTable dt = new DataTable();
        Boolean check_click = false;
        int countFirst = 0;
        int typeChart = 0;

        private Word.Paragraphs wordparagraphs;
        private Word.Paragraph wordparagraph;
        public Form1()
        {
            InitializeComponent();

        }

        //метод для обработки неправильно введенных данных
        private bool checkError()
        {

            for (int i = 0; i < countFirst; i++)
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
            if (countFirst != 0)
            {
                //очистка графика
                foreach (var series in chart2.Series)
                {
                    series.Points.Clear();
                }

                while (dataGridView1.Rows.Count > countFirst)
                {
                    dataGridView1.Rows.RemoveAt(Convert.ToInt32(dataGridView1.Rows.Count - 1));
                }

                if (checkError())
                {
                    //парс входных значений
                    for (int i = 0; i < countFirst; i++)
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

                    //сумма столбцов 
                    double  sumX = 0,
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
                    dataGridView1.Rows[dataGridView1.Rows.Count - 1].HeaderCell.Value = "Сумма";

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
                    double errKoefKorr = Math.Round(Math.Abs(koef_korr / Math.Sqrt((1 - Math.Round(koef_korr * koef_korr, 2)) / (countFirst - 2))), 2);
                    dt.Rows.Add(errKoefKorr);
                    dataGridView1.Rows[dataGridView1.Rows.Count - 1].HeaderCell.Value = "Критерий \r\nдоверенности,\r\nt";

                    //метод для получения Критерия Стюдента
                    //____________________________________ какой уровень значимости выбрать не знаю (выбрал как в лекциях - 0,05)
                    Chart chart1 = new Chart();
                    double koef_znach = 0.05;
                    double t_tabl = Math.Round(chart1.DataManipulator.Statistics.InverseTDistribution(koef_znach, countFirst - 2), 2);
                    dt.Rows.Add(t_tabl, koef_znach);
                    dataGridView1.Rows[dataGridView1.Rows.Count - 1].HeaderCell.Value = "Критерий \r\nСтьюдента,\r\nпри коэффициенте значимости";

                    //делаем вывод после сравнения ошибки с Стюдента
                    string vivod_korr = "Коэфф ошибки корреляции и Стюдента совпадают";

                    if (errKoefKorr > t_tabl)
                    {
                        vivod_korr = "Связь присутствует";
                    }
                    if (errKoefKorr < t_tabl)
                    {
                        vivod_korr = "Связь случайна";
                    }
                    dt.Rows.Add(vivod_korr);
                    dataGridView1.Rows[dataGridView1.Rows.Count - 1].HeaderCell.Value = "Вывод \r\nо значимости коэффициента корреляции";

                    //рисование графика
                    double[] X = new double[countFirst];
                    double[] Y = new double[countFirst];

                    //поиск коэфф регрессии
                    double k = Math.Round((countFirst * sumXY - sumX * sumY) / (countFirst * sumX2 - sumX * sumX), 4);
                    double b = Math.Round(avgY - k * avgX, 2);


                    for (int i = 0; i < countFirst; i++)
                    {
                        //парсим значения Х и У
                        X[i] = Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value);
                        Y[i] = Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value);

                        //добавляем точки и линию на график
                        chart2.Series["Series1"].Points.AddXY(X[i], Y[i]);
                        //если выбрана линейная ф-я
                        if (typeChart == 0) chart2.Series["line"].Points.AddXY(X[i], k * X[i] + b);

                        //степенная
                        if (typeChart == 1) chart2.Series["line"].Points.AddXY(X[i], b * Math.Pow(X[i],k) );

                        //если выбрана гипербола
                        if (typeChart == 2)
                        {
                            chart2.Series["line"].ChartType = SeriesChartType.Spline;
                            chart2.Series["line"].Points.AddXY(X[i], b + k * (1 / X[i]) );
                        }
                        //показательная
                        if (typeChart == 3) chart2.Series["line"].Points.AddXY(X[i], b*Math.Pow(k,X[i]));

                        //если выбрана логарифмическая ф-я
                        if (typeChart == 4) chart2.Series["line"].Points.AddXY(X[i], b + k*Math.Log(X[i]));


                    }
                    //если выбрана линейная ф-я
                    if (typeChart == 0) chart2.Series["line"].LegendText = "y=" + k + "*X+" + b;

                    //степенная
                    if (typeChart == 1) chart2.Series["line"].LegendText = "y=" + k + "*" + b + "^x";
                    //если выбрана гипербола
                    if (typeChart == 2) chart2.Series["line"].LegendText = "y=" + b + "+" + k + "*(1/X)";
                    //показательная
                    if (typeChart == 3) chart2.Series["line"].LegendText = "y=" + b + "*" + k + "^x";

                    //если выбрана логарифмическая ф-я
                    if (typeChart == 4) chart2.Series["line"].LegendText = "y="+b+"+"+k+"*Ln(x)";

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

            //выбор значения по умолчанию
            comboBox1.SelectedItem = "Линейная ф-я";
            comboBox1.SelectedIndexChanged += ComboBox1_SelectedIndexChanged;


        }

        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            typeChart = comboBox1.SelectedIndex;
            button1_Click(sender, e);
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

                //количество исходных данных
                countFirst = dataGridView1.Rows.Count;
                checkError();
            }      
        }

        private void экспортToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (countFirst != 0)
            {
                Stream stream;
                saveFileDialog1.Filter = "txt files (*.txt)|*.txt|CSV files (*.csv)|*.csv|All files (*.*)|*.*";
                saveFileDialog1.FileName = "";
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    if ((stream = saveFileDialog1.OpenFile()) != null)
                    {
                        StreamWriter streamWriter = new StreamWriter(stream, System.Text.Encoding.Default);
                        try
                        {
                            streamWriter.Write(";\t");
                            for (int i=0;i<dt.Columns.Count;i++)
                            {
                                streamWriter.Write(Convert.ToString(dt.Columns[i].ColumnName)+";\t");
                            }
                            streamWriter.WriteLine();

                            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                            {
                                streamWriter.Write(Convert.ToString(dataGridView1.Rows[i].HeaderCell.Value).Replace("\r\n"," ")+ ";\t");
                                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                                {
                                   streamWriter.Write(dataGridView1.Rows[i].Cells[j].Value.ToString() + ";\t");
                                }
                                streamWriter.WriteLine();
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                        
                        streamWriter.Close();
                        stream.Close();

                    }
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (chart2.Series != null)
            {
                saveFileDialog2.Filter = "JPG files (*.jpg)|*.jpg|PNG files (*.png)|*.png";
                if (saveFileDialog2.ShowDialog() == DialogResult.OK)
                {
                    string path = saveFileDialog2.FileName;
                    switch (saveFileDialog2.FilterIndex)
                    {
                        case 1: chart2.SaveImage(path, System.Windows.Forms.DataVisualization.Charting.ChartImageFormat.Jpeg); break;
                        case 2: chart2.SaveImage(path, System.Windows.Forms.DataVisualization.Charting.ChartImageFormat.Png);break;
                    }

                
                }
                saveFileDialog2.FileName = "";
            }
        }

        public void Export_Data_To_Word(DataGridView DGV, string filename)
        {
            if (DGV.Rows.Count != 0)
            {
             
                int RowCount = DGV.Rows.Count;
                int ColumnCount = DGV.Columns.Count+1;
                Object[,] DataArray = new object[RowCount, ColumnCount];

                //заполняем массив
                int r = 0;
                for (int c = 1; c <= ColumnCount - 1; c++)
                {
                    for (r = 0; r <= RowCount - 1; r++)
                    {
                        DataArray[r, c] = DGV.Rows[r].Cells[c-1].Value;
                    }
                } 

                //создаем объект
                Word.Document oDoc = new Word.Document();

                //открытие приложения ворда
                oDoc.Application.Visible = true;

                //выбираем ориентацию страницы
                oDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientPortrait;
                dynamic oRange = oDoc.Content.Application.Selection.Range;
               
                string oTemp = "";
               
                for (r = 0; r <= RowCount - 1; r++)
                {
                    for (int c = 0; c <= ColumnCount - 1; c++)
                    {
                        oTemp = oTemp + DataArray[r, c] + "\t";

                    }
                }

                //протабулированный текст
                oRange.Text = oTemp;
                //указываем тип разделителя - ТАБ'ы
                object Separator = Word.WdTableFieldSeparator.wdSeparateByTabs;
                object ApplyBorders = true;
                object AutoFit = true;
                object AutoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitContent;

                //преобразует текст  в таблицу
                oRange.ConvertToTable(ref Separator, ref RowCount, ref ColumnCount,
                                      Type.Missing, Type.Missing, ref ApplyBorders,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, ref AutoFit, ref AutoFitBehavior, Type.Missing);

                oRange.Select();

                oDoc.Application.Selection.Tables[1].Select();

                //настройки шрифта для таблицы
                oDoc.Application.Selection.Tables[1].Range.Font.Name = "Times New Roman";
                oDoc.Application.Selection.Tables[1].Range.Font.Size = 14;

                //добавляем обводку таблицы
                oDoc.Application.Selection.Tables[1].Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                oDoc.Application.Selection.Tables[1].Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

               //текст не может быть разделен по разрыву страницы
                oDoc.Application.Selection.Tables[1].Rows.AllowBreakAcrossPages = 0;
                //выравнивание по левому краю
                oDoc.Application.Selection.Tables[1].Rows.Alignment = 0;

                oDoc.Application.Selection.Tables[1].Rows[1].Select();

                //отступ после заголовков столбца в строках
                oDoc.Application.Selection.InsertRowsAbove(1);
                oDoc.Application.Selection.Tables[1].Rows[1].Select();

                //стиль строк заголовка
                //полужирный 
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Bold = 1;

                //добавляем заголовки столбцов
                for (int c = 0; c <= ColumnCount - 2; c++)
                {
                    oDoc.Application.Selection.Tables[1].Cell(1, c + 2).Range.Text = DGV.Columns[c].HeaderText;
                }

                //добавляем заголовки строк
                for (int c = 0; c <= RowCount - 1; c++)
                {
                    oDoc.Application.Selection.Tables[1].Cell(c + 2, 0).Range.Text = DGV.Rows[c].HeaderCell.Value.ToString();
                }

                //сохранение
                oDoc.SaveAs2(filename);

            }
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();

            sfd.Filter = "Word Documents (*.docx)|*.docx";

            sfd.FileName = "export.docx";

            if (sfd.ShowDialog() == DialogResult.OK)
            {

                Export_Data_To_Word(dataGridView1, sfd.FileName);
            }
        }
    }
}
