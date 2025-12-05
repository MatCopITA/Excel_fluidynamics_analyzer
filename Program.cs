using DocumentFormat.OpenXml.Drawing;
using OfficeOpenXml;
using System;
using System.ComponentModel;
using System.IO;

public class Program
{
  static string GetNextOutputFilename(string folderPath, string inputFilename)
  {
    string baseName = System.IO.Path.GetFileNameWithoutExtension(inputFilename);
    int counter = 1;
    string fileDir;

    do
    {
      fileDir = System.IO.Path.Combine(folderPath, $"{baseName}_output_{counter}.xlsx");
      counter++;
    }
    while (File.Exists(fileDir));

    return fileDir;
  }

  static string EnsureOutputFolder(string inputFolder)
  {
    string outputFolder = System.IO.Path.Combine(inputFolder, "Output");

    if (!Directory.Exists(outputFolder))
    {
      Directory.CreateDirectory(outputFolder);
      Console.WriteLine($"Created folder: {outputFolder}");
    }
    return outputFolder;
  }

  static void Main(string[] args)
  {
    ExcelPackage.License.SetNonCommercialPersonal("MyName");

    string title = "- Excel Data Analyzer -";
    Console.WriteLine(String.Format("{0," + ((Console.WindowWidth / 2) + (title.Length / 2)) + "}", title));

    string? path = "";
    string extension = "*.xlsx";
    string filePath = "";
    string[] files;

    Console.Write($"Insert the folder path (copy/paste from file explorer): ");
    path = Console.ReadLine();

    if (string.IsNullOrEmpty(path) || !Directory.Exists(path))
    {
      Console.WriteLine("Invalid Option. Press [Enter] to continue...");
      if (Console.ReadKey().Key == ConsoleKey.Enter)
      {
        Console.Clear();
        Main(args);
      }
    }
    else
    { 
     Console.WriteLine($"\nListing all {extension} files in directory: {path}\n");
     files = Directory.GetFiles(path, extension);

    if (files.Length == 0)
    {
      Console.WriteLine("File not found.");
      return;
    }

    for (int i = 0; i < files.Length; i++) Console.WriteLine($"{i + 1}) {System.IO.Path.GetFileName(files[i])}  -->  {files[i]}");

    Console.Write($"\nSelect a file from the list (Put only the number in the front): {Environment.NewLine}");

    if (int.TryParse(Console.ReadLine(), out int choose) &&
        choose >= 1 && choose <= files.Length)
    {
      filePath = files[choose - 1];

      Console.WriteLine($"\nFile chosen: {System.IO.Path.GetFileName(filePath)}");
      Console.WriteLine($"Complete path: {filePath}");
    }
    else Console.WriteLine("Invalid Option.");

    if (filePath == null)
    {
      Console.WriteLine("File not found.");
      return;
    }

      using (var package = new ExcelPackage(new FileInfo(filePath)))
      {
        var worksheet = package.Workbook.Worksheets[0];
        int rowCount = worksheet.Dimension.Rows;

        worksheet.Cells[1, 8].Value = "Re";
        worksheet.Cells[1, 9].Value = "f";
        worksheet.Cells[1, 10].Value = "Δp (Pa)";
        worksheet.Cells[1, 11].Value = "P (W)";
        worksheet.Cells[1, 12].Value = "F_drv (N)";
        worksheet.Cells[1, 13].Value = "l_v (m^2/s^2)";
        worksheet.Cells[1, 14].Value = "τ_w (Pa)";
        worksheet.Cells[1, 15].Value = "Q (m³/s)";

        for (int row = 2; row <= rowCount; row++)
        {
          double v = Convert.ToDouble(worksheet.Cells[row, 1].Value ?? 0);
          double D = Convert.ToDouble(worksheet.Cells[row, 2].Value ?? 0);
          double K = Convert.ToDouble(worksheet.Cells[row, 3].Value ?? 0);
          double rho = Convert.ToDouble(worksheet.Cells[row, 4].Value ?? 0);
          double mu = Convert.ToDouble(worksheet.Cells[row, 5].Value ?? 0);
          double L = Convert.ToDouble(worksheet.Cells[row, 6].Value ?? 0);
          double Q = Convert.ToDouble(worksheet.Cells[row, 7].Value ?? 0);

          if (v == 0 && Q > 0) v = (4 * Q) / (Math.PI * Math.Pow(D, 2));

          K = K * 1e-6;

          double Re = (rho * v * D) / mu;
          double f;

          if (K == 0)
          {
            if (Re < 2100)
              f = 16 / Re;
            else
              f = 0.079 * Math.Pow(Re, -0.25);
          }
          else
          {
            if (Re < 2100)
            {
              f = 16 / Re;
            }
            else
            {
              double f_local = 1e-10;
              double f_old;
              double maxIter = 1e10;
              double tolerance = 1e-10;

              for (int i = 0; i < maxIter; i++)
              {
                f_old = f_local;

                f_local = 1.0 / Math.Pow((-1.7 * Math.Log((K / D) + (4.67 / (Re * Math.Sqrt(f_local)))) + 2.28), 2); 

                if (f_local <= 0) throw new InvalidOperationException("Non-physical friction factor computed.");

                if (Math.Abs(f_local - f_old) < tolerance) break;
              }
              f = f_local;
            }
          }

          double Q_calc;

          if (Q == 0) Q_calc = ((Math.PI * Math.Pow(D, 2)) / 4) * v;
          else Q_calc = Q;

          double p_drop = (L > 0) ? (f * (L / D) * (rho * v * v / 2)) : 0;
          double Pow = p_drop * (Math.PI * D * D / 4) * v;
          double F_drv = rho * v * (Math.PI * D * D / 4);
          double l_v = f * (L / D) * (v * v / (2 * Math.PI));
          double tau_w = (f * rho * v * v) / 8;

          worksheet.Cells[row, 8].Value = Re;
          worksheet.Cells[row, 9].Value = f;
          worksheet.Cells[row, 10].Value = p_drop;
          worksheet.Cells[row, 11].Value = Pow;
          worksheet.Cells[row, 12].Value = F_drv;
          worksheet.Cells[row, 13].Value = l_v;
          worksheet.Cells[row, 14].Value = tau_w;
          worksheet.Cells[row, 15].Value = Q_calc;
        }

        string? inputFolder = System.IO.Path.GetDirectoryName(filePath);
        if (string.IsNullOrEmpty(inputFolder))
          throw new InvalidOperationException("Could not determine input folder.");
        string outputFolder = EnsureOutputFolder(inputFolder);
        string outputFile = GetNextOutputFilename(outputFolder, System.IO.Path.GetFileName(filePath));
        worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
        package.SaveAs(new FileInfo(outputFile));

        Console.WriteLine($"Output saved as: {outputFile}");

        Console.WriteLine($"{Environment.NewLine}Press any key to continue or press [Esc] to exit...{Environment.NewLine}");
        if (Console.ReadKey().Key == ConsoleKey.Escape) return;
        else
        {
          Console.Clear();
          Main(args);
        }
      }
    }
  }
}

