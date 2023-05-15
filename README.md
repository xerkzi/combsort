using System;
using System.IO;
using OfficeOpenXml;


class Program
{
    static void Main(string[] args)
    {
     
        try
        {
            static Tuple<int, int> CombSort(int[] arr, int n)
            {
                int gap = n;
                bool swapped = true;
                int swaps = 0; 
                int comparisons = 0; 
                while (gap > 1 || swapped)
                {
                    if (gap > 1)
                    {
                        gap = (int)(gap / 1.247330950103979);
                    }
                    int i = 0;
                    swapped = false;
                    while (i + gap < n)
                    {
                        comparisons++;
                        if (arr[i] > arr[i + gap])
                        {
                            int temp = arr[i];
                            arr[i] = arr[i + gap];
                            arr[i + gap] = temp;
                            swapped = true;
                            swaps++; 
                        }
                        i++;
                    }
                }
                Console.WriteLine("\nОтсортированный массив:");
                Console.WriteLine("---------------------");
                Console.WriteLine("| Индекс | Значение |");
                Console.WriteLine("---------------------");
                for (int i = 0; i < arr.Length; i++)
                {
                    Console.WriteLine($"| {i,6} | {arr[i],8} |");
                }
                Console.WriteLine("---------------------");
                return Tuple.Create(swaps, comparisons);
            }
            static Tuple<int, int> CombSortReverse(int[] arr, int n)
            {
                int gap = n;
                bool swapped = true;
                int swaps = 0; 
                int comparisons = 0; 
                while (gap > 1 || swapped)
                {
                    comparisons++;
                    gap = (int)(gap / 1.247330950103979);
                    if (gap < 1)
                        gap = 1;

                    swapped = false;
                    for (int i = 0; i + gap < n; i++)
                    {
                        if (arr[i] < arr[i + gap])
                        {
                            int temp = arr[i];
                            arr[i] = arr[i + gap];
                            arr[i + gap] = temp;
                            swapped = true;
                            swaps++; 
                        }
                    }
                }
                Console.WriteLine("\nОтсортированный массив:");
                Console.WriteLine("---------------------");
                Console.WriteLine("| Индекс | Значение |");
                Console.WriteLine("---------------------");
                for (int i = 0; i < arr.Length; i++)
                {
                    Console.WriteLine($"| {i,6} | {arr[i],8} |");
                }
                Console.WriteLine("---------------------");

                return Tuple.Create(swaps, comparisons);
            }
            Console.WriteLine("Введите размер массива:");
            int n = Convert.ToInt32(Console.ReadLine()); 
            int[] arr = new int[n]; 

            Console.WriteLine("Выберите способ заполнения массива:");
            Console.WriteLine("1 - Автоматически");
            Console.WriteLine("2 - Вручную");
            int choice = Convert.ToInt32(Console.ReadLine());

            if (choice == 1)
            {
                Console.WriteLine("Введите диапазон ОТ:");
                int min = Convert.ToInt32(Console.ReadLine());
                Console.WriteLine("Введите диапазон ДО:");
                int max = Convert.ToInt32(Console.ReadLine());
                if (min > max)
                {
                    Console.WriteLine("Ошибка: диапазон ОТ должен быть меньше или равен диапазону ДО.");
                    return;
                }
             
                Random rnd = new Random();
                for (int i = 0; i < n; i++)
                {
                    int num = rnd.Next(min, max + 1);
                    arr[i] = num;
                }
            }
            else if (choice == 2)
            {
                // заполняем массив вручную
                Console.WriteLine("Введите элементы массива:");
                for (int i = 0; i < n; i++)
                {
                    arr[i] = Convert.ToInt32(Console.ReadLine());
                }
            }
            else
            {
                Console.WriteLine("Некорректный выбор");
                return;
            }

            Console.WriteLine("Массив:");
            foreach (int num in arr)
            {
                Console.Write(num + " ");
            }
            int[] arrold = new int[n]; 
            for (int i = 0; i < arr.Length; i++)
            {
                arrold[i] = arr[i];
            }
            Console.WriteLine();

            int swaps;
            int comparisons;

            Console.WriteLine("Выберите способ сортировки массива:");
            Console.WriteLine("1 - По возрастанию");
            Console.WriteLine("2 - По убыванию");
            Console.WriteLine("3 - Случайным образом");
            int choice1 = Convert.ToInt32(Console.ReadLine());
            if (choice1 == 1)
            {
                System.Diagnostics.Stopwatch stopwatch = new System.Diagnostics.Stopwatch();
                stopwatch.Start();
                Tuple<int, int> results = CombSort(arr, n);
                swaps = results.Item1;
                comparisons = results.Item2;
                Console.WriteLine("\nКоличество сравнений: " + comparisons);
                Console.WriteLine("Количество перестановок: " + swaps);
                stopwatch.Stop();
                Console.WriteLine("Время выполнения: " + stopwatch.ElapsedMilliseconds + " мс");

            }
            else if (choice1 == 2)
            {
                System.Diagnostics.Stopwatch stopwatch = new System.Diagnostics.Stopwatch();
                stopwatch.Start();
                Tuple<int, int> results = CombSortReverse(arr, n);
                swaps = results.Item1;
                comparisons = results.Item2;
                Console.WriteLine("\nКоличество сравнений: " + comparisons);
                Console.WriteLine("Количество перестановок: " + swaps);
                stopwatch.Stop();
                Console.WriteLine("Время выполнения: " + stopwatch.ElapsedMilliseconds + " мс");
            }
            else if(choice1 == 3)
            {
                Random rand = new Random();
                int randomNum = rand.Next(0, 2);

                if (randomNum == 1)
                {
                    System.Diagnostics.Stopwatch stopwatch = new System.Diagnostics.Stopwatch();
                    stopwatch.Start();
                    Tuple<int, int> results = CombSort(arr, n);
                    swaps = results.Item1;
                    comparisons = results.Item2;
                    Console.WriteLine("Массив будет отсортирован по возрастанию");
                    Console.WriteLine("\nКоличество сравнений: " + comparisons);
                    Console.WriteLine("Количество перестановок: " + swaps);
                    stopwatch.Stop();
                    Console.WriteLine("Время выполнения: " + stopwatch.ElapsedMilliseconds + " мс");
                }
                else
                {
                    Console.WriteLine("Массив будет отсортирован по убыванию");
                    System.Diagnostics.Stopwatch stopwatch = new System.Diagnostics.Stopwatch();
                    stopwatch.Start();
                    Tuple<int, int> results = CombSortReverse(arr, n);
                    swaps = results.Item1;
                    comparisons = results.Item2;
                    Console.WriteLine("\nКоличество сравнений: " + comparisons);
                    Console.WriteLine("Количество перестановок: " + swaps);
                    stopwatch.Stop();
                    Console.WriteLine("Время выполнения: " + stopwatch.ElapsedMilliseconds + " мс");
                }
                return;
            }
            else
            {
                Console.WriteLine("Некорректный выбор");
                return;
            }

            Console.WriteLine("Для сохранения данных в файл Excel нажмите 1 \n Для выхода из программы нажмите любую клавишу");
            int choice2= Convert.ToInt32(Console.ReadLine());
            if (choice2 == 1)
            {
                Console.WriteLine("Сохранение данных...");
                FileInfo file = new FileInfo(@"D:\1\combsort.xlsx");
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (ExcelPackage package = new ExcelPackage(file))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Сортировка массива");
                    worksheet.Cells.EntireColumn.AutoFit();
                    worksheet.Row(1).Height = 25;
                    worksheet.Column(1).Width = 25;
                    worksheet.Row(2).Height = 25;
                    worksheet.Column(2).Width = 25;
                    worksheet.Row(3).Height = 25;
                    worksheet.Column(3).Width = 25;
                    worksheet.Row(4).Height = 25;
                    worksheet.Column(4).Width = 25;
                    worksheet.Cells[1, 1].Value = "Массив до сортировки:";
                    worksheet.Cells[1, 2].Value = "Массив после сортировки:";
                    worksheet.Cells[1, 3].Value = "Количество сравнений:";
                    worksheet.Cells[1, 4].Value = "Количество перестановок:";
                    worksheet.Cells[2, 3].Value = comparisons;
                    worksheet.Cells[2, 4].Value = swaps;

                    for (int i = 0; i < arrold.Length; i++)
                    {
                        worksheet.Cells[i + 2, 1].Value = arrold[i];
                    }
                    for (int i = 0; i < arr.Length; i++)
                    {
                        worksheet.Cells[i + 2, 2].Value = arr[i];
                    }
                    package.Save();
                }
                Console.WriteLine("Данные успешно сохранены в файл combsort.xlsx");
            }
            else
            {
                Console.ReadKey();
            }
              
        }
            
        catch
        {
            try
            {
                Console.WriteLine("Ошибка выполнения программы. Перезапустить программу?");
                Console.WriteLine("1 - Да");
                Console.WriteLine("2 - Нет");
                int choice3 = Convert.ToInt32(Console.ReadLine());
                if (choice3 == 1)
                {
                    Main(args);
                }
                else if (choice3 == 2)
                {
                    Environment.Exit(0);
                }
                else
                {
                    Console.WriteLine("Некорректный выбор");
                    return;
                }
            }
            catch
            {
                Console.WriteLine("Ошибка");
            }
        }
    }
}


