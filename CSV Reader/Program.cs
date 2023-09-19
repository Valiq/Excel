
using System.Text.RegularExpressions;
using System;
using System.Threading.Tasks.Dataflow;
using System.IO;

internal class Program
{

    private static void Main(string[] args)
    {
        string file = "error.csv";
        string newFile = "result2.csv";
        string errorFile = "error2.csv";
        string path = "E:\\Desktop\\";

        List<string[]> list = new List<string[]>();
        List<string> finalList = new List<string>();
        List<string> errorList = new List<string>();

        string[] lines = File.ReadAllLines($"{path}{file}");
        Console.CursorVisible = false;
        string[] nameRows = { "ID", "WHOLESALE_NET_PRICE", "GROSS_NET_PRICE", 
                            "VAT", "INDEX", "CORE_ID",
                            "CATALOG_ID", "NEW_ID", "MANUFACTURER_ID", 
                            "MANUFACTURER", "GROUP_ID", "GROUP" };

        string nameRow = "";
        foreach (var row in nameRows)
        {
            nameRow = nameRow + row + ";";
        }

        nameRow = nameRow.TrimEnd(';');

        finalList.Add(nameRow);

        int k = 0; int er = 0;
        foreach (var line in lines)
        {
            string newRow = "";

            string[] rows = line.Split(';');

            if (rows.Count() == 11)
            {
                for (int i = 0; i < nameRows.Count(); i++)
                {
                    if (i == 7)
                    {
                        if (rows[6].Count() < 3)
                        {
                            newRow = newRow + rows[4] + ";";
                        }
                        else
                            newRow = newRow + rows[6] + ";";
                    }
                    else if (i > 7)
                    {
                        newRow = newRow + rows[i - 1] + ";";
                    }
                    else
                    {
                        newRow = newRow + rows[i] + ";";
                    }
                }

                newRow = newRow.TrimEnd(';');

                if (k > 0)
                {
                    finalList.Add(newRow);
                }
            }
            else
            {
                er++;
                string error = "";
                foreach (var row in rows)
                    error = error + row + ";";
                
                error = error.TrimEnd(';');

                errorList.Add(error);
            }

            k++;

            Console.SetCursorPosition(0, 0);
            Console.WriteLine($"Количество обработанных строк {k}");
            Console.WriteLine($"С ошибками {er}");
        }

        Console.WriteLine("Обработка завершена\nЗапись в файл без ошибок\n");

        using (StreamWriter writer = new StreamWriter($"{path}{newFile}", false))
        {
            foreach (var row in finalList)
                 writer.WriteLine(row);
        }

        Console.WriteLine("Запись в файл ошибок");

        using (StreamWriter writer = new StreamWriter($"{path}{errorFile}", false))
        {
            foreach (var row in errorList)
                writer.WriteLine(row);
        }

        Console.WriteLine("END");
    }
}