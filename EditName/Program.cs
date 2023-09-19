using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
internal class Program
{
    static Excel.Application excel = new Excel.Application();

    private static void Main(string[] args)
    {
        Console.CursorVisible = false;
        AppDomain.CurrentDomain.ProcessExit += new EventHandler(CurrentDomain_ProcessExit);

        string pathDir = "E:\\Desktop\\Stellox\\Test";
        string[] files = Directory.GetFiles(pathDir);

        int fileCounter = 0;

        StreamWriter write = new StreamWriter($"{pathDir}\\result.txt");

        foreach (var fileName in files)
        {
            if (fileName.Contains("result.txt"))
            {
                continue;
            }
            string[] str = fileName.Split('\\');
            fileCounter++;
            Console.WriteLine($"{fileCounter}. {str[str.Length - 1]}");
            write.WriteLine($"{fileCounter}. {str[str.Length - 1]}");

            excel.Workbooks.Open(fileName);

            try
            {
                (int Left, int Top) value = Console.GetCursorPosition();
                int left = value.Left;
                int top = value.Top;

                int i = 4;
                List<string> resultList = new List<string>();

                while (!string.IsNullOrEmpty(excel.Sheets[5].Cells[i, 2].Value))
                {
                    Console.WriteLine($"Количество обработанных строк {i - 4}");
                    string buffer = excel.Sheets[5].Cells[i, 3].Value.ToString();
                    buffer.Trim().Replace('-',' ');

                    string newStr = "";

                    for (int j = 0; j < buffer.Length; j++)
                    {
                        if ((j < buffer.Length - 1) && (j > 0))
                        {
                            if ((buffer[j] == ' ') && (buffer[j + 1] == ' '))
                            {
                                continue;
                            }

                            var c1 = buffer[j];
                            var c2 = buffer[j + 1];

                            if ((Regex.IsMatch(buffer[j].ToString(), "[а-я]|[a-z]") && Regex.IsMatch(buffer[j + 1].ToString(), "[А-Я]|[A-Z]")) ||
                                 (Regex.IsMatch(buffer[j - 1].ToString(), "[А-Я]|[A-Z]|[а-я]|[a-z]") && Regex.IsMatch(buffer[j].ToString(), "[.]") && Regex.IsMatch(buffer[j + 1].ToString(), "[А-Я]|[A-Z]|[а-я]|[a-z]")))
                            {
                                newStr += $"{buffer[j]} ";
                            }
                            else
                            {
                                newStr += $"{buffer[j]}";
                            }
                        }
                        else
                        {
                            newStr += $"{buffer[j]}";
                        }
                    }

                    string[] cuts = newStr.Trim().Split(' ');

                    newStr = "";

                    bool flag = true;
                    foreach (var cut in cuts)
                    {
                        if (Regex.IsMatch(cut, "^[a-z]+$", RegexOptions.IgnoreCase) && (cut.Length > 1) && (!cut.Contains("STELLOX")) && (!cut.Contains("ABS")) && flag)
                        {
                            newStr += $"для {cut} ";
                            flag = false;
                        }
                        else
                        {
                            newStr += $"{cut.Trim()} ";
                        }
                    }

                    string[] cutSplesh = newStr.Split('/');

                    if (cutSplesh.Length > 4)
                    {
                        newStr = $"{cutSplesh[0]}/{cutSplesh[1]}/{cutSplesh[2]}/{cutSplesh[cutSplesh.Length - 1]}";
                    }

                    excel.Sheets[5].Cells[i, 3] = newStr.TrimEnd(' ');

                    List<string> tempList = new List<string>();

                    foreach (var cut in cuts)
                    {
                        if (Regex.IsMatch(cut, "[а-я]-[а-я]"))
                        {
                            tempList.Add(cut);
                        }

                        if (Regex.IsMatch(cut, "[а-я][.]$"))
                        {
                            tempList.Add(cut);
                        }

                        if ((cut.Length < 5) && Regex.IsMatch(cut, "[а-я]"))
                        {
                            tempList.Add(cut);
                        }
                    }

                    var result = resultList.Union(tempList);

                    resultList = result.ToList();

                    i++;
                    Console.SetCursorPosition(left, top);
                }

                foreach (var row in resultList)
                {
                    Console.WriteLine(row);
                    write.WriteLine(row);
                }

                Console.WriteLine("\n");
                write.WriteLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                KillExcel(excel);
                write.Close();
            }

            excel.ActiveWorkbook.Close(true);
        }

        write.Close();
    }

    [DllImport("User32.dll")]
    public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int ProcessId);
    private static void KillExcel(Excel.Application theApp)
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
        if (excel is not null)
        {
            KillExcel(excel);
        }
        Console.WriteLine(" exit");
    }
}