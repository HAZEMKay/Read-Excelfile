



using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

class Program
{
    static void Main()
    {
        string filePath = "C:\\Users\\hazem\\Desktop\\Algorithm\\Students.xlsx";
        string sortedFilePath = "C:\\Users\\hazem\\Desktop\\Algorithm\\SortedStudents.xlsx"; //new file
        List<(string name, int degree)> studentInfo = ReadDegreesAndNamesFromExcel(filePath);

        PrintDegreesAndNames(studentInfo, "Unsorted Degrees and Names");

        List<int> degrees = studentInfo.Select(info => info.degree).ToList();
        RadixSort(degrees);

        for (int i = 0; i < studentInfo.Count; i++)
        {
            studentInfo[i] = (studentInfo[i].name, degrees[i]);
        }

        PrintDegreesAndNames(studentInfo, "Sorted Degrees and Names");

        CreateExcelFile(studentInfo, sortedFilePath);

        Console.ReadLine();
    }

    static List<(string name, int degree)> ReadDegreesAndNamesFromExcel(string filePath)
    {
        List<(string name, int degree)> studentInfo = new List<(string name, int degree)>();

        FileInfo file = new FileInfo(filePath);
        using (ExcelPackage package = new ExcelPackage(file))
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

            int rowCount = worksheet.Dimension.Rows;

            for (int i = 2; i <= rowCount; i++)
            {
                string cellValue = worksheet.Cells[i, 1].Value?.ToString();
                string[] parts = cellValue?.Split(' ');

                if (parts != null && parts.Length > 1)
                {
                    string name = parts[0];
                    string numericPart = parts[1];
                    string[] numericParts = numericPart.Split('/');

                    if (numericParts.Length > 1 && int.TryParse(numericParts[1], out int degree))
                    {
                        studentInfo.Add((name, degree));
                    }
                    else
                    {
                        Console.WriteLine($"Invalid data at row {i}: {cellValue}");
                    }
                }
                else
                {
                    Console.WriteLine($"Invalid data at row {i}: {cellValue}");
                }
            }
        }

        return studentInfo;
    }

    static void RadixSort(List<int> arr)
    {
        int max = GetMax(arr);

        for (int exp = 1; max / exp > 0; exp *= 10)
        {
            CountSort(arr, exp);
        }
    }

    static int GetMax(List<int> arr)
    {
        int max = arr[0];
        for (int i = 1; i < arr.Count; i++)
        {
            if (arr[i] > max)
            {
                max = arr[i];
            }
        }
        return max;
    }

    static void CountSort(List<int> arr, int exp)
    {
        int[] output = new int[arr.Count];
        int[] count = new int[10];

        for (int i = 0; i < arr.Count; i++)
        {
            count[(arr[i] / exp) % 10]++;
        }

        for (int i = 1; i < 10; i++)
        {
            count[i] += count[i - 1];
        }

        for (int i = arr.Count - 1; i >= 0; i--)
        {
            output[count[(arr[i] / exp) % 10] - 1] = arr[i];
            count[(arr[i] / exp) % 10]--;
        }

        for (int i = 0; i < arr.Count; i++)
        {
            arr[i] = output[i];
        }
    }

    static void PrintDegreesAndNames(List<(string name, int degree)> studentInfo, string label)
    {
        Console.WriteLine(label + ":");
        foreach (var info in studentInfo)
        {
            Console.WriteLine($"{info.name}: {info.degree}");
        }
        Console.WriteLine();
    }

    static void CreateExcelFile(List<(string name, int degree)> studentInfo, string filePath)
    {
        FileInfo file = new FileInfo(filePath);
        using (ExcelPackage package = new ExcelPackage(file))
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sorted Students");
            worksheet.Cells[1, 1].Value = "Name";
            worksheet.Cells[1, 2].Value = "Degree";

            for (int i = 0; i < studentInfo.Count; i++)
            {
                worksheet.Cells[i + 2, 1].Value = studentInfo[i].name;
                worksheet.Cells[i + 2, 2].Value = studentInfo[i].degree;
            }

            package.Save();
        }
    }
}
