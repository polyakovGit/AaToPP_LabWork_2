using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace lab2
{
    class Program
    {
        delegate void ArrayFillOptions(double[,] arr);

        static void ArrayFillUL(double[,] arr)
        {
            for (int i = 0; i < arr.GetLength(0); ++i)
            {
                for (int j = 0; j < arr.GetLength(1); ++j)
                    arr[i, j] = i / (j + 1.0);
            }
        }

        static void ArrayFillUR(double[,] arr)
        {
            for (int i = 0; i < arr.GetLength(0); ++i)
            {
                for (int j = arr.GetLength(1) - 1; j >= 0; --j)
                    arr[i, j] = i / (j + 1.0);
            }
        }

        static void ArrayFillDL(double[,] arr)
        {
            for (int i = arr.GetLength(0) - 1; i >= 0; --i)
            {
                for (int j = 0; j < arr.GetLength(1); ++j)
                    arr[i, j] = i / (j + 1.0);
            }
        }

        static void ArrayFillDR(double[,] arr)
        {
            for (int i = arr.GetLength(0) - 1; i >= 0; --i)
            {
                for (int j = arr.GetLength(1) - 1; j >= 0; --j)
                    arr[i, j] = i / (j + 1.0);
            }
        }

        static void ArrayFillLU(double[,] arr)
        {
            for (int j = 0; j < arr.GetLength(1); ++j)
            {
                for (int i = 0; i < arr.GetLength(0); ++i)
                    arr[i, j] = i / (j + 1.0);
            }
        }
        static void ArrayFillRU(double[,] arr)
        {
            for (int j = arr.GetLength(1) - 1; j >= 0; --j)
            {
                for (int i = 0; i < arr.GetLength(0); ++i)
                    arr[i, j] = i / (j + 1.0);
            }
        }

        static void ArrayFillLD(double[,] arr)
        {
            for (int j = 0; j < arr.GetLength(1); ++j)
            {
                for (int i = arr.GetLength(0) - 1; i >= 0; --i)
                    arr[i, j] = i / (j + 1.0);
            }
        }

        static void ArrayFillRD(double[,] arr)
        {
            for (int j = arr.GetLength(1) - 1; j >= 0; --j)
            {
                for (int i = arr.GetLength(0) - 1; i >= 0; --i)
                    arr[i, j] = i / (j + 1.0);
            }
        }
        static void fillArray(double[,] arr)
        {
            for (int i = 0; i < arr.GetLength(0); ++i)
                for (int j = 0; j < arr.GetLength(1); ++j)
                    arr[i, j] = i / (j + 1.0);//rand.Next() % 100;
        }


        static void printArray(double[,] arr)
        {
            for (int i = 0; i < arr.GetLength(0); ++i)
            {
                for (int j = 0; j < arr.GetLength(1); ++j)
                    Console.Write($"{arr[i, j],3}");
                Console.WriteLine();
            }
        }
        class ArrayFill
        {
            public ArrayFillOptions fillMethod;
            public string name;
        }
        static void Main(string[] args)
        {
            if (Debugger.IsAttached)
            {
                Console.WriteLine("Warning, debugger attached!");
            }

            ResultsToExcel();

            //ArrayFill[] fillOptions = {
            //    new ArrayFill{fillMethod=ArrayFillUL,name="UL"},
            //    new ArrayFill{fillMethod=ArrayFillUR,name="UR"},
            //    new ArrayFill{fillMethod=ArrayFillDL,name="DL"},
            //    new ArrayFill{fillMethod=ArrayFillDR,name="DR"},
            //    new ArrayFill{fillMethod=ArrayFillLU,name="LU"},
            //    new ArrayFill{fillMethod=ArrayFillRU,name="RU"},
            //    new ArrayFill{fillMethod=ArrayFillLD,name="LD"},
            //    new ArrayFill{fillMethod=ArrayFillRD,name="RD"},
            //};
            ////ArrayFillOptions[] fillOptions ={ArrayFillUL,ArrayFillUR,ArrayFillDL,ArrayFillDR,
            ////    ArrayFillLU,ArrayFillRU,ArrayFillLD,ArrayFillRD};
            //const int N = 5000;
            //double[,] arr = new double[N, N];

            //Stopwatch sw = new Stopwatch();

            //sw.Restart();
            //fillArray(arr);
            //sw.Stop();
            //Console.WriteLine(sw.ElapsedMilliseconds);

            //long[] time = new long[100];
            //double t = 1.9842;//t-коэффициент доверия по таблице функции Лапласа если n=100 и 95% уровень значимости

            //for (int i = 0; i < fillOptions.Length; ++i)
            //{
            //    double avg = 0, D = 0, S;
            //    //sw = new Stopwatch();
            //    for (int k = 0; k < 100; ++k)
            //    {
            //        sw.Restart();
            //        //fillOptions[i].Invoke(arr);
            //        fillOptions[i].fillMethod.Invoke(arr);
            //        sw.Stop();
            //        time[k] = sw.ElapsedMilliseconds;
            //        avg += time[k];
            //        //Console.WriteLine(sw.ElapsedMilliseconds);
            //    }
            //    avg /= time.Length;//среднее значение выборки
            //    for (int j = 0; j < time.Length; ++j)
            //    {
            //        D += (time[j] - avg) * (time[j] - avg);
            //        //Console.WriteLine(D);
            //    }
            //    D /= time.Length;//дисперсия выборки
            //    S = Math.Sqrt(D);//среднеквадратическое отклонение
            //    Console.WriteLine($"интервал для {fillOptions[i].name} " +
            //        $"({avg - t * S / Math.Sqrt(time.Length):f2}; {avg + t * S / Math.Sqrt(time.Length):f2})");
            //    Console.WriteLine($"среднее {avg:f2} дисперсия {D:f2} среднеквадратическое {S:f2}");
            //}

        }
        static void ResultsToExcel(/*long [] array*/)
        {
            //Array.Sort(array);
            //long min = array[0], q1 = array[24], median = array[49], q3 = array[74], max = array[99];
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWb = xlApp.Workbooks.Open(Directory.GetCurrentDirectory()+ "/OutputTextFile.xlsx");//ошибка при открытии книги
            Excel.Worksheet xlSht = xlWb.Sheets[1];
            int iLastRow = xlSht.Cells[xlSht.Rows.Count, "A"].End[Excel.XlDirection.xlUp].Row;  //последняя заполненная строка в столбце А
            for (int i = 1; i < 51; i++)
            {
                iLastRow++;
                xlSht.Cells[iLastRow, "A"].Value = i.ToString();
            }
            xlApp.Visible=true;
            xlWb.Close(true);
            xlApp.Quit();
        }
    }
}
