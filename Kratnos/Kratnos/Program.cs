
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;


internal class Program
{
    static Excel.Application excel = new Excel.Application();

    private static void Main(string[] args)
    {
            string pathPattern = "C:\\Users\\Admin\\Downloads\\Telegram Desktop\\Ozon.xlsx";
            string pathDir = "C:\\Users\\Admin\\Desktop\\kratTest";        

            string[] listArt = new string[2965];
            string[] listVal = new string[2965];

        AppDomain.CurrentDomain.ProcessExit += new EventHandler(CurrentDomain_ProcessExit);

        excel.Workbooks.Open(pathPattern);

        try
        {
            Console.CursorVisible = false;
            Console.WriteLine("Массив ещё не сформирован !!! \n");

            for (int i = 0; i < listArt.Length; i++)
            {
                Console.SetCursorPosition(0, 0);
                Console.WriteLine("Формирование массива...          \n");
                Console.WriteLine($"Добавлено {i+1} элементов из {listArt.Length} \n");

                Excel.Range? rng = excel.Cells[i + 2, 1] as Excel.Range;
                listArt[i] = rng.Value2;

                rng = excel.Cells[i + 2, 7] as Excel.Range;
                listVal[i] = rng.Value.ToString();
            }

            excel.ActiveWorkbook.Close(true);

            string[] files = Directory.GetFiles(pathDir);

            int fileCounter = 0;

            foreach (var fileName in files)
            {
                string[] str = fileName.Split('\\'); 

                try
                {
                    fileCounter++;
                    Console.WriteLine($"{fileCounter}. {str[str.Length - 1]}");
                    excel.Workbooks.Open(fileName);

                    int indexCol = 0; bool flag = false;
                    for (int i = 1; i < 80; i++)
                    {
                        Excel.Range? rng = excel.Sheets[5].Cells[2, i] as Excel.Range;
                        if (rng?.Value != null)
                        {
                            if (rng.Value.Contains("Кратность покупки"))
                            {
                                Console.WriteLine($"\t{rng?.Value2.ToString()}");

                                indexCol = i;
                                flag = true;

                                Console.WriteLine("\tИндекс колонки " + indexCol);
                            }
                        }
                    }

                    if (!flag)
                    {
                        Console.WriteLine("\tКарточка без кратности (Удаление файла)\n");
                        excel.ActiveWorkbook.Close(true);
                        File.Delete(fileName);

                        //excel.ActiveWorkbook.SaveAs("(без кратности) " + fileName);
                        //excel.ActiveWorkbook.SaveAs(fileName.Remove(fileName.Length - 5) + "(без кратности).xlsx");
                    }


                    if (flag)
                    {
                        int indexRow = 4;
                        int countEdit = 0; int countNull = 0; int countRow = 0;
                        Excel.Range? rng = excel.Sheets[5].Cells[indexRow, 2] as Excel.Range;

                        (int Left, int Top) value = Console.GetCursorPosition();
                        int left = value.Left;
                        int top = value.Top;

                        while (rng?.Value != null)
                        {
                            int index = Array.IndexOf(listArt, rng?.Value);

                            if (index != -1)
                            {
                                excel.Sheets[5].Cells[indexRow, indexCol].Value = listVal[index].ToString();
                                countEdit++;
                            }

                            if (index == -1)
                            {
                                if (excel.Sheets[5].Cells[indexRow, indexCol].Value == null)
                                {
                                    countNull++;
                                    Excel.Range rngDel = (Excel.Range)excel.Sheets[5].Rows[indexRow, Type.Missing];
                                    //Console.WriteLine(rngDel.Rows.Count.ToString());
                                    rngDel.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                                    indexRow--;
                                }
                            }

                            countRow++;
                            indexRow++;
                            rng = excel.Sheets[5].Cells[indexRow, 2] as Excel.Range;

                            Console.SetCursorPosition(left, top);

                            Console.WriteLine($"\tОбработано  {countRow} строк");
                            Console.WriteLine($"\tИзменено {countEdit} записей");
                            Console.WriteLine($"\t{countNull} Удалено строк (пустые ячейки) \n");

                        }

                        excel.ActiveWorkbook.Close(true);

                        if (countRow == countNull)
                        {
                            Console.WriteLine("   Удаление пустого файла \n");
                            File.Delete(fileName);
                        }

                        //excel.ActiveWorkbook.SaveAs("(изменено) " + fileName, Excel.XlSaveAsAccessMode.xlShared);
                    }

                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Ошибка в файле {str[str.Length - 1]} \n");
                    Console.WriteLine(ex + "\n");
                }
            }

            Console.WriteLine("\t !!! ГОТОВО !!!");
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex);
            KillExcel(excel);
        }
    }


    [DllImport("User32.dll")]
    public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int ProcessId);
    private static void KillExcel(Application theApp)
    {
        int id = 0;
        IntPtr intptr = new IntPtr(theApp.Hwnd);
        System.Diagnostics.Process p = null;
        try
        {
            GetWindowThreadProcessId(intptr, out id);
            p = System.Diagnostics.Process.GetProcessById(id);
            if (p != null)
            {
                p.Kill();
                p.Dispose();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(" KillExcel:" + ex.Message);
        }
    }

    static void CurrentDomain_ProcessExit(object sender, EventArgs e)
    {
        KillExcel(excel);
        Console.WriteLine(" exit");
    }
}