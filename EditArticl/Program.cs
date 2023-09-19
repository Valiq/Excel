using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

internal class Program
{
    static Excel.Application excel = new Excel.Application();

    private static void Main(string[] args)
    {
        //string pathPattern = "C:\\Users\\Admin\\Downloads\\Telegram Desktop\\Ozon.xlsx";
        string pathDir = "E:\\Desktop\\Start";

        AppDomain.CurrentDomain.ProcessExit += new EventHandler(CurrentDomain_ProcessExit);

        string[] files = Directory.GetFiles(pathDir);

        int fileCounter = 0;

        foreach (var fileName in files)
        {
            string[] strName = fileName.Split('\\');

            try
            {
                fileCounter++;
                Console.WriteLine($"{fileCounter}. {strName[strName.Length - 1]}");
                excel.Workbooks.Open(fileName);

                int indexCol = 0; bool flag = false;
                for (int i = 1; i < 80; i++)
                {
                    Excel.Range? rng = excel.Sheets[5].Cells[2, i] as Excel.Range;
                    if (rng?.Value != null)
                    {
                        if (rng.Value.Contains("Партномер (артикул производителя)"))
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
                    Console.WriteLine("\tОтсутсвие колонки с Партномерами\n");
                    excel.ActiveWorkbook.Close(true);
                    File.Delete(fileName);
                }

                if (flag)
                {
                    int indexRow = 4;
                    int countDel = 0; int countRow = 0;
                    Excel.Range? rng = excel.Sheets[5].Cells[indexRow, 2] as Excel.Range;

                    (int Left, int Top) value = Console.GetCursorPosition();
                    int left = value.Left;
                    int top = value.Top;

                    while (rng?.Value != null)
                    {
                        Excel.Range? partNumber = excel.Sheets[5].Cells[indexRow, indexCol] as Excel.Range;

                        string artDef = rng?.Value.ToString();
                        string artProd = partNumber?.Value.ToString();

                        //if (partNumber?.Value != null)
                        //if (rng?.Value.ToString()[0] != '0')
                        if (artDef.Split('=')[0].Equals(artProd))
                        {
                            countDel++;
                            Excel.Range rngDel = (Excel.Range)excel.Sheets[5].Rows[indexRow, Type.Missing];
                            rngDel.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                            indexRow--;
                        }
                        else
                        {
                            partNumber.NumberFormat = "@";
                            partNumber.Value = artDef.Split('=')[0];
                        }

                        countRow++;
                        indexRow++;
                        rng = excel.Sheets[5].Cells[indexRow, 2] as Excel.Range;

                        Console.SetCursorPosition(left, top);

                        Console.WriteLine($"\tОбработано  {countRow} строк");
                        Console.WriteLine($"\t{countDel} Удалено строк (имеют значение) \n");
                    }

                    excel.ActiveWorkbook.Close(true);

                    if (countRow == countDel)
                    {
                        Console.WriteLine("   Удаление пустого файла \n");
                        File.Delete(fileName);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка в файле {strName[strName.Length - 1]} \n");
                Console.WriteLine(ex + "\n");
            }
        }

        Console.WriteLine("\t !!! ГОТОВО !!!");
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