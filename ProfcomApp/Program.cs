using System.Text;

using OfficeOpenXml;

namespace ProfcomApp;

public class Program
{
    private static void Main()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        Console.InputEncoding = Encoding.UTF8;

        Console.WriteLine("Hello, please write the path to the \"answers\" file from the Google form :) \n");
        var inputFormResponsesFile = Console.ReadLine();

        var inputExcelFile = new FileInfo(inputFormResponsesFile);
        if (!inputExcelFile.Exists)
        {
            Console.WriteLine("File not found.");
            return;
        }

        Console.WriteLine("Please write the path to the folder where the generated files will be \n");
        var outputFolder = Console.ReadLine();

        using (var inputPackage = new ExcelPackage(inputExcelFile))
        {
            var inputWorksheet = inputPackage.Workbook.Worksheets[0];
            FormatBuilder.FormatColumns(inputWorksheet);
            inputPackage.Save();
        }

        using (var inputPackage = new ExcelPackage(inputExcelFile))
        {
            var inputWorksheet = inputPackage.Workbook.Worksheets[0];
            for (var course = 1; course <= 4; course++)
            {
                var outputFile = Path.Combine(outputFolder, $"{course} курс (2020-2024).xlsx");
                var outputFileInfo = new FileInfo(outputFile);

                using var outputPackage = new ExcelPackage(outputFileInfo);
                var outputRowIndex = 1;

                var uniqueGroupNames = new HashSet<string>();
                var mainInformationPage = outputPackage.Workbook.Worksheets.Add($"{course} курс");

                for (var row = 2; !string.IsNullOrEmpty(inputWorksheet.Cells[row, 1].Text); row++)
                {
                    if (inputWorksheet.Cells[row, 4].Text == course.ToString())
                        inputWorksheet.Cells[row, 1, row, inputWorksheet.Dimension.End.Column]
                            .Copy(mainInformationPage.Cells[outputRowIndex++, 1]);
                    else
                        continue;

                    var groupName = inputWorksheet.Cells[row, 5].Text.Trim();

                    if (!uniqueGroupNames.Contains(groupName))
                    {
                        var groupWorksheet = outputPackage.Workbook.Worksheets.Add(groupName);
                        FormatBuilder.CreateHeader(groupWorksheet);
                        uniqueGroupNames.Add(groupName);
                    }
                }

                var sortRange = mainInformationPage.Cells[1, 1, mainInformationPage.Dimension.End.Row,
                    mainInformationPage.Dimension.End.Column];
                mainInformationPage.Cells[sortRange.ToString()].Sort(x => x.SortBy.Column(1));
                outputPackage.Save();

                var rowCounters = new Dictionary<string, int>();

                for (var row = 1; !string.IsNullOrEmpty(mainInformationPage.Cells[row, 1].Text); row++)
                {
                    var groupName = mainInformationPage.Cells[row, 5].Text;
                    var groupWorksheet = outputPackage.Workbook.Worksheets[groupName];

                    rowCounters.TryAdd(groupName, 3);

                    var rowCounter = rowCounters[groupName];

                    FormatBuilder.FillColumns(groupWorksheet, mainInformationPage, rowCounter, row);

                    rowCounters[groupName]++;
                }

                var commonPage = outputPackage.Workbook.Worksheets.Add("Общий");
                FormatBuilder.CreateHeader(commonPage);

                foreach (var ugn in uniqueGroupNames)
                {
                    var groupWorksheet = outputPackage.Workbook.Worksheets[ugn];
                    FormatBuilder.FillingCommonInformation(groupWorksheet, commonPage);
                    FormatBuilder.CreateTableBorders(groupWorksheet);
                }

                outputPackage.Save();
            }
        }

        Console.WriteLine("Ready!");
    }
}