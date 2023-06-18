using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using WinForms = System.Windows.Forms;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using MessageBox = System.Windows.MessageBox;
using Aspose.Cells;
using Microsoft.Office.Interop.Excel;
using System.Drawing;
using Style = Aspose.Cells.Style;
using Workbook = Aspose.Cells.Workbook;
using Worksheet = Aspose.Cells.Worksheet;
using Window = System.Windows.Window;

namespace Cnopochci
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void UploadBtn_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog openFileDialog = new OpenFileDialog();

            bool? resp = openFileDialog.ShowDialog();

            if (resp == true)
            {
                string filepath = openFileDialog.FileName;

                try
                {
                    string pr = filepath.Split('.')[1];
                    if ((pr == "xls") || (pr == "xlsx"))
                    {
                        TextBl.Text = filepath;
                    }
                    else
                    {
                        MessageBox.Show("Файл находится в неверном формате");
                    }
                }
                catch 
                {
                    MessageBox.Show("Пожалуйста, выберите файл");
                }
            }
        }

        private void UploadBtn_Click2(object sender, RoutedEventArgs e)
        {
            WinForms.FolderBrowserDialog dialog = new WinForms.FolderBrowserDialog();
            dialog.ShowDialog();
            TextBl2.Text = dialog.SelectedPath;
        }

        string ToMyChar(int val)
        {
            string znach = "";
            if (val < 91)
            {
                znach += Convert.ToChar(val);
            }
            else
            {
                val = val - 64;
                int a = val / 26;
                int b = val % 26;
                znach += Convert.ToChar(a + 64);
                znach += Convert.ToChar(b + 64);
            }
            return znach;
        }

        private void Preobraz_Click(object sender, RoutedEventArgs e)
        {
            if ((TextBl.Text == " ") || (TextBl2.Text == " "))
            {
                MessageBox.Show("Пожалуйста, введите путь к файлу и папку, в которую будет сохранён файл");
            }
            else
            {
                try
                {
                    LoadOptions loadOptions3 = new LoadOptions(LoadFormat.SpreadsheetML);
                    Workbook workbook = new Workbook(TextBl.Text, loadOptions3);

                    Worksheet DannProecta = workbook.Worksheets[0];
                    Worksheet VxodDann = workbook.Worksheets[1];

                    int i = workbook.Worksheets.Add();
                    Worksheet ObemPoStad = workbook.Worksheets[i];
                    ObemPoStad.Name = "1. Объем по стад вода и нефть";
                    ObemPoStad.Cells.StandardWidth = 15;

                    i = workbook.Worksheets.Add();
                    Worksheet PritokPoVodePersent = workbook.Worksheets[i];
                    PritokPoVodePersent.Name = " 2. Приток по воде %";
                    PritokPoVodePersent.Cells.StandardWidth = 15;

                    i = workbook.Worksheets.Add();
                    Worksheet PritokPoNeftiPersent = workbook.Worksheets[i];
                    PritokPoNeftiPersent.Name = "3. Приток по нефти %";
                    PritokPoNeftiPersent.Cells.StandardWidth = 15;

                    i = workbook.Worksheets.Add();
                    Worksheet PritokPoVodeABC = workbook.Worksheets[i];
                    PritokPoVodeABC.Name = "4. Приток по воде абс";
                    PritokPoVodeABC.Cells.StandardWidth = 15;

                    i = workbook.Worksheets.Add();
                    Worksheet PritokPoNeftiABC = workbook.Worksheets[i];
                    PritokPoNeftiABC.Name = "5. Приток по нефти абс";
                    PritokPoNeftiABC.Cells.StandardWidth = 15;

                    Style style = workbook.CreateStyle();
                    style.HorizontalAlignment = TextAlignmentType.Center;
                    style.VerticalAlignment = TextAlignmentType.Center;

                    Style styleData = workbook.CreateStyle();
                    styleData.Number = 22;

                    Style styleSumma = workbook.CreateStyle();
                    styleSumma.Number = 10;

                    Style styleDouble = workbook.CreateStyle();
                    styleDouble.Number = 2;

                    Style newDat = workbook.CreateStyle();
                    newDat.Number = 14;

                    int ColStr = DannProecta.Cells["E5"].IntValue;
                    int kol = 0;

                    ObemPoStad.Cells["B4"].PutValue("дата, время");
                    ObemPoStad.Cells["B4"].SetStyle(style);
                    ObemPoStad.Cells[ToMyChar(66 + ColStr + 3) + "4"].PutValue("дата, время");
                    ObemPoStad.Cells[ToMyChar(66 + ColStr + 3) + "4"].SetStyle(style);
                    ObemPoStad.Cells.Merge(1, 1, 1, ColStr + 3);
                    ObemPoStad.Cells["B2"].PutValue("Объем по стадиям, нефть");
                    ObemPoStad.Cells["B2"].SetStyle(style);
                    ObemPoStad.Cells["B3"].SetStyle(style);
                    ObemPoStad.Cells.Merge(1, 4 + ColStr, 1, ColStr + 3);
                    ObemPoStad.Cells[ToMyChar(70 + ColStr) + "2"].PutValue("Объем по стадиям, вода");
                    ObemPoStad.Cells[ToMyChar(70 + ColStr) + "2"].SetStyle(style);
                    ObemPoStad.Cells.Merge(2, 2, 1, ColStr);
                    ObemPoStad.Cells["C3"].PutValue("Количество нефтерастворимого трассера");
                    ObemPoStad.Cells["C3"].SetStyle(style);
                    ObemPoStad.Cells.Merge(2, 5 + ColStr, 1, ColStr);
                    ObemPoStad.Cells[ToMyChar(67 + ColStr + 3) + "3"].PutValue("Количество водорастворимого трассера");
                    ObemPoStad.Cells[ToMyChar(67 + ColStr + 3) + "3"].SetStyle(style);
                    ObemPoStad.Cells.Merge(2, 2 + ColStr, 2, 1);
                    ObemPoStad.Cells[ToMyChar(67 + ColStr) + "3"].PutValue("Дебит\n нефти, т");
                    ObemPoStad.Cells[ToMyChar(67 + ColStr) + "3"].SetStyle(style);
                    ObemPoStad.Cells.Merge(2, 3 + ColStr, 2, 1);
                    ObemPoStad.Cells[ToMyChar(67 + ColStr + 1) + "3"].PutValue("Сумма\n трасеров");
                    ObemPoStad.Cells[ToMyChar(67 + ColStr + 1) + "3"].SetStyle(style);
                    ObemPoStad.Cells[ToMyChar(69 + ColStr) + "2"].PutValue("Объем по стадиям, вода");
                    ObemPoStad.Cells[ToMyChar(69 + ColStr) + "2"].SetStyle(style);
                    ObemPoStad.Cells.Merge(2, 5 + 2 * ColStr, 2, 1);
                    ObemPoStad.Cells[ToMyChar(70 + 2 * ColStr) + "3"].PutValue("Дебит воды, м3");
                    ObemPoStad.Cells[ToMyChar(70 + 2 * ColStr) + "3"].SetStyle(style);
                    ObemPoStad.Cells.Merge(2, 6 + 2 * ColStr, 2, 1);
                    ObemPoStad.Cells[ToMyChar(71 + 2 * ColStr) + "3"].PutValue("Сумма трасеров");
                    ObemPoStad.Cells[ToMyChar(71 + 2 * ColStr) + "3"].SetStyle(style);

                    //H1, H2, ...
                    for (int j = 0; j < ColStr; j++)
                    {
                        ObemPoStad.Cells[ToMyChar(67 + j) + "4"].PutValue(VxodDann.Cells[ToMyChar(67 + j) + "5"].StringValue);
                        ObemPoStad.Cells[ToMyChar(67 + j) + "4"].SetStyle(style);
                        ObemPoStad.Cells[ToMyChar(70 + ColStr + j) + "4"].PutValue(VxodDann.Cells[ToMyChar(68 + ColStr + j) + "5"].StringValue);
                        ObemPoStad.Cells[ToMyChar(70 + ColStr + j) + "4"].SetStyle(style);
                        PritokPoVodePersent.Cells[ToMyChar(67 + j) + "4"].PutValue(VxodDann.Cells[ToMyChar(68 + ColStr + j) + "5"].StringValue);
                        PritokPoVodePersent.Cells[ToMyChar(67 + j) + "4"].SetStyle(style);
                        PritokPoNeftiPersent.Cells[ToMyChar(67 + j) + "4"].PutValue(VxodDann.Cells[ToMyChar(67 + j) + "5"].StringValue);
                        PritokPoNeftiPersent.Cells[ToMyChar(67 + j) + "4"].SetStyle(style);
                        PritokPoVodeABC.Cells[ToMyChar(67 + j) + "4"].PutValue(VxodDann.Cells[ToMyChar(68 + ColStr + j) + "5"].StringValue);
                        PritokPoVodeABC.Cells[ToMyChar(67 + j) + "4"].SetStyle(style);
                        PritokPoNeftiABC.Cells[ToMyChar(67 + j) + "4"].PutValue(VxodDann.Cells[ToMyChar(67 + j) + "5"].StringValue);
                        PritokPoNeftiABC.Cells[ToMyChar(67 + j) + "4"].SetStyle(style);
                    }

                    // date and time
                    int l = 0;
                    while (VxodDann.Cells["B" + (l + 6).ToString()].StringValue != "")
                    {
                        ObemPoStad.Cells["B" + (l + 5).ToString()].Formula = "='Входные данные'!B" + (l + 6).ToString();
                        ObemPoStad.Cells["B" + (l + 5).ToString()].SetStyle(styleData);
                        ObemPoStad.Cells[ToMyChar(69 + ColStr) + (l + 5).ToString()].Formula = "='Входные данные'!B" + (l + 6).ToString();
                        ObemPoStad.Cells[ToMyChar(69 + ColStr) + (l + 5).ToString()].SetStyle(styleData);
                        PritokPoVodePersent.Cells["A" + (l + 5).ToString()].Formula = "='Входные данные'!B" + (l + 6).ToString();
                        PritokPoVodePersent.Cells["A" + (l + 5).ToString()].SetStyle(styleData);
                        PritokPoNeftiPersent.Cells["A" + (l + 5).ToString()].Formula = "='Входные данные'!B" + (l + 6).ToString();
                        PritokPoNeftiPersent.Cells["A" + (l + 5).ToString()].SetStyle(styleData);
                        PritokPoVodeABC.Cells["A" + (l + 5).ToString()].Formula = "='Входные данные'!B" + (l + 6).ToString();
                        PritokPoVodeABC.Cells["A" + (l + 5).ToString()].SetStyle(styleData);
                        PritokPoNeftiABC.Cells["A" + (l + 5).ToString()].Formula = "='Входные данные'!B" + (l + 6).ToString();
                        PritokPoNeftiABC.Cells["A" + (l + 5).ToString()].SetStyle(styleData);
                        l += 1;
                    }

                    //summ of trasssers and debit nefti
                    for (int j = 0; j < l; j++)
                    {
                        double summNeft = 0, summVoda = 0;
                        for (int g = 0; g < ColStr; g++)
                        {
                            summNeft += VxodDann.Cells[ToMyChar(67 + g) + (j + 6).ToString()].DoubleValue;
                            summVoda += VxodDann.Cells[ToMyChar(68 + ColStr + g) + (j + 6).ToString()].DoubleValue;
                        }
                        ObemPoStad.Cells[ToMyChar(68 + ColStr) + (j + 5).ToString()].PutValue(summNeft);
                        ObemPoStad.Cells[ToMyChar(67 + ColStr) + (j + 5).ToString()].Formula = $"='Входные данные'!{ToMyChar(67 + ColStr) + (j + 6).ToString()}";
                        PritokPoNeftiPersent.Cells[ToMyChar(67 + ColStr) + (j + 5).ToString()].Formula = $"='Входные данные'!{ToMyChar(67 + ColStr) + (j + 6).ToString()}";
                        ObemPoStad.Cells[ToMyChar(71 + 2 * ColStr) + (j + 5).ToString()].PutValue(summVoda);
                        ObemPoStad.Cells[ToMyChar(70 + 2 * ColStr) + (j + 5).ToString()].Formula = $"='Входные данные'!{ToMyChar(68 + 2 * ColStr) + (j + 6).ToString()}";
                        PritokPoVodePersent.Cells[ToMyChar(67 + ColStr) + (j + 5).ToString()].Formula = $"='Входные данные'!{ToMyChar(68 + 2 * ColStr) + (j + 6).ToString()}";
                        PritokPoVodeABC.Cells[ToMyChar(67 + ColStr) + (j + 5).ToString()].Formula = $"='Входные данные'!{ToMyChar(68 + 2 * ColStr) + (j + 6).ToString()}";
                        PritokPoNeftiABC.Cells[ToMyChar(67 + ColStr) + (j + 5).ToString()].Formula = $"='Входные данные'!{ToMyChar(67 + ColStr) + (j + 6).ToString()}";

                    }

                    //kolichestvo trassera
                    for (int j = 0; j < l; j++)
                    {
                        int NomStroki = 66;
                        for (int h = 0; h < ColStr; h++)
                        {
                            NomStroki++;
                            ObemPoStad.Cells[ToMyChar(NomStroki) + (j + 5).ToString()].Formula = $"='Входные данные'!{ToMyChar(NomStroki) + (j + 6).ToString()}/'1. Объем по стад вода и нефть'!{ToMyChar(68 + ColStr) + (j + 5).ToString()}";
                            ObemPoStad.Cells[ToMyChar(NomStroki) + (j + 5).ToString()].SetStyle(styleSumma);
                            ObemPoStad.Cells[ToMyChar(NomStroki + 3 + ColStr) + (j + 5).ToString()].Formula = $"='Входные данные'!{ToMyChar(NomStroki + 1 + ColStr) + (j + 6).ToString()}/'1. Объем по стад вода и нефть'!{ToMyChar(71 + 2 * ColStr) + (j + 5).ToString()}";
                            ObemPoStad.Cells[ToMyChar(NomStroki + 3 + ColStr) + (j + 5).ToString()].SetStyle(styleSumma);
                        }
                    }

                    PritokPoVodePersent.Cells.Merge(1, 0, 1, ColStr + 2);
                    PritokPoVodePersent.Cells["A2"].PutValue("Объем по стадиям, вода");
                    PritokPoVodePersent.Cells["A2"].SetStyle(style);
                    PritokPoVodePersent.Cells.Merge(2, 0, 2, 1);
                    PritokPoVodePersent.Cells["A3"].PutValue("дата, время");
                    PritokPoVodePersent.Cells["A3"].SetStyle(style);
                    PritokPoVodePersent.Cells.Merge(2, ColStr + 2, 2, 1);
                    PritokPoVodePersent.Cells[ToMyChar(67 + ColStr) + "3"].PutValue("Дебит воды, м3");
                    PritokPoVodePersent.Cells[ToMyChar(67 + ColStr) + "3"].SetStyle(style);
                    PritokPoVodePersent.Cells.Merge(2, 1, 1, ColStr);
                    PritokPoVodePersent.Cells["B3"].PutValue("Количество водорастворимого трассера");
                    PritokPoVodePersent.Cells["B3"].SetStyle(style);

                    for (int j = 0; j < l; j++)
                    {
                        PritokPoVodePersent.Cells["B" + (j + 5).ToString()].PutValue(j + 1);
                        int NomStroki = 66;
                        for (int h = 0; h < ColStr; h++)
                        {
                            PritokPoVodePersent.Cells[ToMyChar(NomStroki + 1) + (j + 5).ToString()].Formula = $"='1. Объем по стад вода и нефть'!{ToMyChar(NomStroki + 4 + ColStr) + (j + 5).ToString()}";
                            PritokPoVodePersent.Cells[ToMyChar(NomStroki + 1) + (j + 5).ToString()].SetStyle(styleSumma);
                            NomStroki++;
                        }
                    }
                    PritokPoVodePersent.Cells.HideColumn(1);

                    int chartIndex = PritokPoVodePersent.Charts.Add(Aspose.Cells.Charts.ChartType.AreaStacked, 2, ColStr + 4, 20, 21);
                    Aspose.Cells.Charts.Chart chart = PritokPoVodePersent.Charts[chartIndex];
                    chart.SetChartDataRange($"B4:{ToMyChar(66 + ColStr) + (4 + l).ToString()}", true);

                    PritokPoNeftiPersent.Cells.Merge(1, 0, 1, ColStr + 2);
                    PritokPoNeftiPersent.Cells["A2"].PutValue("Объем по стадиям, нефть");
                    PritokPoNeftiPersent.Cells["A2"].SetStyle(style);
                    PritokPoNeftiPersent.Cells.Merge(2, 0, 2, 1);
                    PritokPoNeftiPersent.Cells["A3"].PutValue("дата, время");
                    PritokPoNeftiPersent.Cells["A3"].SetStyle(style);
                    PritokPoNeftiPersent.Cells.Merge(2, ColStr + 2, 2, 1);
                    PritokPoNeftiPersent.Cells[ToMyChar(67 + ColStr) + "3"].PutValue("Дебит нефти, т");
                    PritokPoNeftiPersent.Cells[ToMyChar(67 + ColStr) + "3"].SetStyle(style);
                    PritokPoNeftiPersent.Cells.Merge(2, 1, 1, ColStr);
                    PritokPoNeftiPersent.Cells["B3"].PutValue("Количество нефтерастворимого трассера");
                    PritokPoNeftiPersent.Cells["B3"].SetStyle(style);

                    for (int j = 0; j < l; j++)
                    {
                        PritokPoNeftiPersent.Cells["B" + (j + 5).ToString()].PutValue(j + 1);
                        int NomStroki = 66;
                        for (int h = 0; h < ColStr; h++)
                        {
                            PritokPoNeftiPersent.Cells[ToMyChar(NomStroki + 1) + (j + 5).ToString()].Formula = $"='1. Объем по стад вода и нефть'!{ToMyChar(NomStroki + 1) + (j + 5).ToString()}";
                            PritokPoNeftiPersent.Cells[ToMyChar(NomStroki + 1) + (j + 5).ToString()].SetStyle(styleSumma);
                            NomStroki++;
                        }
                    }
                    PritokPoNeftiPersent.Cells.HideColumn(1);

                    chartIndex = PritokPoNeftiPersent.Charts.Add(Aspose.Cells.Charts.ChartType.AreaStacked, 2, ColStr + 4, 20, 21);
                    Aspose.Cells.Charts.Chart chart1 = PritokPoNeftiPersent.Charts[chartIndex];
                    chart1.SetChartDataRange($"B4:{ToMyChar(66 + ColStr) + (4 + l).ToString()}", true);

                    PritokPoVodeABC.Cells.Merge(1, 0, 1, ColStr + 2);
                    PritokPoVodeABC.Cells["A2"].PutValue("Объем по стадиям, вода");
                    PritokPoVodeABC.Cells["A2"].SetStyle(style);
                    PritokPoVodeABC.Cells.Merge(2, 0, 2, 1);
                    PritokPoVodeABC.Cells["A3"].PutValue("дата, время");
                    PritokPoVodeABC.Cells["A3"].SetStyle(style);
                    PritokPoVodeABC.Cells.Merge(2, ColStr + 2, 2, 1);
                    PritokPoVodeABC.Cells[ToMyChar(67 + ColStr) + "3"].PutValue("Дебит воды, м3");
                    PritokPoVodeABC.Cells[ToMyChar(67 + ColStr) + "3"].SetStyle(style);
                    PritokPoVodeABC.Cells.Merge(2, 1, 1, ColStr);
                    PritokPoVodeABC.Cells["B3"].PutValue("Количество водорастворимого трассера");
                    PritokPoVodeABC.Cells["B3"].SetStyle(style);

                    for (int j = 0; j < l; j++)
                    {
                        PritokPoVodeABC.Cells["B" + (j + 5).ToString()].PutValue(j + 1);
                        int NomStroki = 66;
                        for (int h = 0; h < ColStr; h++)
                        {
                            PritokPoVodeABC.Cells[ToMyChar(NomStroki + 1) + (j + 5).ToString()].Formula = $"=' 2. Приток по воде %'!{ToMyChar(NomStroki + 1) + (j + 5).ToString()}*' 2. Приток по воде %'!${ToMyChar(67 + ColStr) + (j + 5).ToString()}";
                            PritokPoVodeABC.Cells[ToMyChar(NomStroki + 1) + (j + 5).ToString()].SetStyle(styleDouble);
                            NomStroki++;
                        }
                    }
                    PritokPoVodeABC.Cells.HideColumn(1);

                    chartIndex = PritokPoVodeABC.Charts.Add(Aspose.Cells.Charts.ChartType.AreaStacked, 2, ColStr + 4, 20, 21);
                    Aspose.Cells.Charts.Chart chart2 = PritokPoVodeABC.Charts[chartIndex];
                    chart2.SetChartDataRange($"B4:{ToMyChar(66 + ColStr) + (4 + l).ToString()}", true);

                    PritokPoNeftiABC.Cells.Merge(1, 0, 1, ColStr + 2);
                    PritokPoNeftiABC.Cells["A2"].PutValue("Объем по стадиям, нефть");
                    PritokPoNeftiABC.Cells["A2"].SetStyle(style);
                    PritokPoNeftiABC.Cells.Merge(2, 0, 2, 1);
                    PritokPoNeftiABC.Cells["A3"].PutValue("дата, время");
                    PritokPoNeftiABC.Cells["A3"].SetStyle(style);
                    PritokPoNeftiABC.Cells.Merge(2, ColStr + 2, 2, 1);
                    PritokPoNeftiABC.Cells[ToMyChar(67 + ColStr) + "3"].PutValue("Дебит нефти, т");
                    PritokPoNeftiABC.Cells[ToMyChar(67 + ColStr) + "3"].SetStyle(style);
                    PritokPoNeftiABC.Cells.Merge(2, 1, 1, ColStr);
                    PritokPoNeftiABC.Cells["B3"].PutValue("Количество нефтерастворимого трассера");
                    PritokPoNeftiABC.Cells["B3"].SetStyle(style);

                    for (int j = 0; j < l; j++)
                    {
                        PritokPoNeftiABC.Cells["B" + (j + 5).ToString()].PutValue(j + 1);
                        int NomStroki = 66;
                        for (int h = 0; h < ColStr; h++)
                        {
                            PritokPoNeftiABC.Cells[ToMyChar(NomStroki + 1) + (j + 5).ToString()].Formula = $"='3. Приток по нефти %'!{ToMyChar(NomStroki + 1) + (j + 5).ToString()}*'3. Приток по нефти %'!${ToMyChar(67 + ColStr) + (j + 5).ToString()}";
                            PritokPoNeftiABC.Cells[ToMyChar(NomStroki + 1) + (j + 5).ToString()].SetStyle(styleDouble);
                            NomStroki++;
                        }
                    }
                    PritokPoNeftiABC.Cells.HideColumn(1);

                    chartIndex = PritokPoNeftiABC.Charts.Add(Aspose.Cells.Charts.ChartType.AreaStacked, 2, ColStr + 4, 20, 21);
                    Aspose.Cells.Charts.Chart chart3 = PritokPoNeftiABC.Charts[chartIndex];
                    chart3.SetChartDataRange($"B4:{ToMyChar(66 + ColStr) + (4 + l).ToString()}", true);

                    String dattim = TextBl3.Text;
                    if (dattim != "")
                    {
                        try
                        {
                            i = workbook.Worksheets.Add();
                            Worksheet DataProfil = workbook.Worksheets[i];
                            DataProfil.Name = "6. Профиль по дате";
                            DataProfil.Cells.StandardWidth = 12;

                            String dat = dattim.Split(' ')[0];
                            String tim = dattim.Split(' ')[1];

                            int date = int.Parse(dat.Split('.')[0]);
                            int month = int.Parse(dat.Split('.')[1]);
                            int year = int.Parse(dat.Split('.')[2]);

                            int hours = int.Parse(tim.Split(':')[0]);
                            int minutes = int.Parse(tim.Split(':')[1]);
                            int seconds = int.Parse(tim.Split(':')[2]);

                            DateTime tt = new DateTime(year, month, date, hours, minutes, seconds);

                            DataProfil.Cells["A2"].PutValue(tt);
                            DataProfil.Cells["A2"].SetStyle(newDat);
                            DataProfil.Cells["A1"].PutValue("Дата");

                            int nomer = -1;
                            for (int j = 0; j < l; j++)
                            {
                                if (VxodDann.Cells["B" + (j + 6).ToString()].DateTimeValue == DataProfil.Cells["A2"].DateTimeValue) nomer = j;
                            }

                            if (nomer == -1)
                            {
                                DataProfil.Cells["C2"].PutValue("Такой даты нет");
                            }
                            else
                            {
                                DataProfil.Cells["C2"].PutValue("Вода");
                                DataProfil.Cells["C3"].PutValue("Нефть");
                                for (int j = 0; j < ColStr; j++)
                                {
                                    DataProfil.Cells[ToMyChar(68 + j) + "1"].PutValue("Порт " + (j + 1).ToString());
                                    DataProfil.Cells[ToMyChar(68 + j) + "2"].Formula = ObemPoStad.Cells[ToMyChar(70 + j + ColStr) + (nomer + 5).ToString()].Formula;
                                    DataProfil.Cells[ToMyChar(68 + j) + "2"].SetStyle(styleSumma);
                                    DataProfil.Cells[ToMyChar(68 + j) + "3"].Formula = ObemPoStad.Cells[ToMyChar(67 + j) + (nomer + 5).ToString()].Formula;
                                    DataProfil.Cells[ToMyChar(68 + j) + "3"].SetStyle(styleSumma);
                                }
                                chartIndex = DataProfil.Charts.Add(Aspose.Cells.Charts.ChartType.Bar, 4, 2, 16, 8);
                                Aspose.Cells.Charts.Chart chart4 = DataProfil.Charts[chartIndex];
                                chart4.SetChartDataRange($"C1:M3", true);
                            }
                        }
                        catch 
                        {
                            MessageBox.Show("Введён неверный формат даты");
                        }
                    }
                    //22.10.2022 16:00:00

                    workbook.Save(TextBl2.Text+"\\Готовое.xls");
                    MessageBox.Show("Данные записаны!");
                }
                catch
                {
                    MessageBox.Show("Ошибка в данных файла");
                }
            }
        }
    }
}