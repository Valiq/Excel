using System;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;


internal class Program
{
    static Excel.Application excel = new Excel.Application();
    private static void Main(string[] args)
    {
        Console.CursorVisible = false;

        string pathDir = "E:\\Desktop\\Yandex";
        AppDomain.CurrentDomain.ProcessExit += new EventHandler(CurrentDomain_ProcessExit);

        string[] files = Directory.GetFiles(pathDir);

        foreach (var fileName in files)
        {
            Console.WriteLine("Начало обработки файла");
            excel.Workbooks.Open(fileName);

            int indexRow = 4; int rowCounter = 0; int editRow = 0;
            Excel.Range? rng = excel.Sheets[2].Cells[indexRow, 3] as Excel.Range;

            (int Left, int Top) value = Console.GetCursorPosition();
            int left = value.Left;
            int top = value.Top;

            while (rng?.Value != null)
            {
                try
                {
                    Excel.Range? partRng = excel.Sheets[2].Cells[indexRow, 13] as Excel.Range;

                    if (partRng?.Value == "")
                    {
                        string[] name = rng.Value.Split('=');
                        excel.Sheets[2].Cells[indexRow, 13].Value = name[0];
                        Console.WriteLine($"\t{name[0]}");
                        editRow++;
                    }
                    //Console.WriteLine($"{rng?.Value}\t\t{partRng?.Value}");

                    indexRow++; rowCounter++;
                    rng = excel.Sheets[2].Cells[indexRow, 3] as Excel.Range;

                    Console.SetCursorPosition(left, top);

                    Console.WriteLine($"\tОбработано  {rowCounter} строк");
                    Console.WriteLine($"\tИзменено {editRow} записей");
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex + "\n");
                }
            }
        }

        excel.ActiveWorkbook.Close(true);
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