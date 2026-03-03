using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;

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


        List<(double Q, double DeltaP, double Re, double f)> dataPoints = new();
        for (int row = 2; row <= rowCount; row++)
        {
          double Q_val = Convert.ToDouble(worksheet.Cells[row, 15].Value ?? 0);
          double dp_val = Convert.ToDouble(worksheet.Cells[row, 10].Value ?? 0);
          double Re_val = Convert.ToDouble(worksheet.Cells[row, 8].Value ?? 0);
          double f_val = Convert.ToDouble(worksheet.Cells[row, 9].Value ?? 0);

          dataPoints.Add((Q_val, dp_val, Re_val, f_val));
        }
        var sortedData = dataPoints
            .OrderBy(d => d.Q)
            .ToList();

        worksheet.Cells[1, 17].Value = "Q_sorted";
        worksheet.Cells[1, 18].Value = "Δp_sorted";
        worksheet.Cells[1, 19].Value = "Re_sorted";
        worksheet.Cells[1, 20].Value = "f_sorted";

        for (int i = 0; i < sortedData.Count; i++)
        {
          worksheet.Cells[i + 2, 17].Value = sortedData[i].Q;
          worksheet.Cells[i + 2, 18].Value = sortedData[i].DeltaP;
          worksheet.Cells[i + 2, 19].Value = sortedData[i].Re;
          worksheet.Cells[i + 2, 20].Value = sortedData[i].f;
        }

        double minRe = double.MaxValue;
        double maxRe = double.MinValue;
        double minF = double.MaxValue;
        double maxF = double.MinValue;
        for (int row = 2; row <= rowCount; row++)
        {
          double Re_val = Convert.ToDouble(worksheet.Cells[row, 8].Value ?? 0);
          double f_val = Convert.ToDouble(worksheet.Cells[row, 9].Value ?? 0);

          if (Re_val > 0 && f_val > 0)
          {
            if (Re_val < minRe) minRe = Re_val;
            if (Re_val > maxRe) maxRe = Re_val;
            if (f_val < minF) minF = f_val;
            if (f_val > maxF) maxF = f_val;
          }
        }

        double marginFactor = 0.1;

        var chart1 = worksheet.Drawings.AddChart("Re_vs_f_log", eChartType.XYScatterLines);
        chart1.Title.Text = "Re vs f";
        chart1.XAxis.LogBase = 10;
        chart1.YAxis.LogBase = 10;
        chart1.XAxis.Title.Text = "Re";
        chart1.YAxis.Title.Text = "f(Re)";
        chart1.SetPosition(22, 0, 0, 0);
        chart1.SetSize(600, 500);
        var series = chart1.Series.Add(
          worksheet.Cells[2, 20, rowCount, 20], //f
          worksheet.Cells[2, 19, rowCount, 19]  //Re
        );
        series.Header = "f vs Re";

        chart1.XAxis.MinorTickMark = eAxisTickMark.None;
        chart1.YAxis.MinorTickMark = eAxisTickMark.None;

        chart1.XAxis.MinValue = minRe * (1 - marginFactor);
        chart1.XAxis.MaxValue = maxRe * (1 + marginFactor);
        chart1.YAxis.MinValue = minF * (1 - marginFactor);
        chart1.YAxis.MaxValue = maxF * (1 + marginFactor);

        var chart2 = worksheet.Drawings.AddChart("Q_vs_Δp", eChartType.XYScatterLines);
        chart2.Title.Text = "Q vs Δp";
        chart2.XAxis.Title.Text = "Q";
        chart2.YAxis.Title.Text = "Δp";
        chart2.SetPosition(22, 0, 12, 0);
        chart2.SetSize(600, 500);
        var series2 = chart2.Series.Add(
          worksheet.Cells[2, 18, rowCount, 18], //Δp
          worksheet.Cells[2, 17, rowCount, 17]  //Q
        );
        series2.Header = "Δp vs Q";

        //worksheet.Column(17).Hidden = true;
        //worksheet.Column(18).Hidden = true;
        //worksheet.Column(19).Hidden = true;
        //worksheet.Column(20).Hidden = true;

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

