using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ProfcomApp;

public static class FormatBuilder
{
    public static void CreateHeader(ExcelWorksheet worksheet)
    {
        using (var headerCells = worksheet.Cells[1, 1, 2, 37])
        {
            headerCells.Style.Fill.PatternType = ExcelFillStyle.Solid;
            headerCells.Style.Fill.BackgroundColor.SetColor(Color.LightYellow);
            headerCells.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
            headerCells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            headerCells.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            headerCells.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            headerCells.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            headerCells.Style.Border.Left.Style = ExcelBorderStyle.Thin;
        }

        worksheet.Cells[1, 11, 1, 14].Merge = true;
        worksheet.Cells[1, 15, 1, 19].Merge = true;
        worksheet.Cells[1, 20, 1, 26].Merge = true;
        worksheet.Cells[1, 28, 1, 33].Merge = true;

        worksheet.Cells[1, 1, 2, 1].Merge = true;
        worksheet.Cells[1, 2, 2, 2].Merge = true;
        worksheet.Cells[1, 3, 2, 3].Merge = true;
        worksheet.Cells[1, 4, 2, 4].Merge = true;
        worksheet.Cells[1, 5, 2, 5].Merge = true;
        worksheet.Cells[1, 6, 2, 6].Merge = true;
        worksheet.Cells[1, 7, 2, 7].Merge = true;
        worksheet.Cells[1, 8, 2, 8].Merge = true;
        worksheet.Cells[1, 9, 2, 9].Merge = true;
        worksheet.Cells[1, 10, 2, 10].Merge = true;
        worksheet.Cells[1, 27, 2, 27].Merge = true;
        worksheet.Cells[1, 34, 2, 34].Merge = true;
        worksheet.Cells[1, 35, 2, 35].Merge = true;
        worksheet.Cells[1, 36, 2, 36].Merge = true;
        worksheet.Cells[1, 37, 2, 37].Merge = true;

        worksheet.Cells[1, 1].Value = "№ пп";
        worksheet.Cells[2, 1].Value = " ";
        worksheet.Cells[1, 2].Value = "ФИО студента";
        worksheet.Cells[2, 2].Value = " ";
        worksheet.Cells[1, 3].Value = "минчане";
        worksheet.Cells[2, 3].Value = " ";
        worksheet.Cells[1, 4].Value = "иногородние студенты";
        worksheet.Cells[2, 4].Value = " ";
        worksheet.Cells[1, 5].Value =
            "имеющие льготы в соответствии с Законом «О социальной защите граждан, пострадавших от катастрофы на Чернобыльской АЭС»";
        worksheet.Cells[2, 5].Value = " ";
        worksheet.Cells[1, 6].Value = "иностранные граждане";
        worksheet.Cells[2, 6].Value = " ";
        worksheet.Cells[1, 7].Value = "студентов платной формы обучения";
        worksheet.Cells[2, 7].Value = " ";
        worksheet.Cells[1, 8].Value = "студентов бюджетной формы обучения";
        worksheet.Cells[2, 8].Value = " ";
        worksheet.Cells[1, 9].Value = "из многодетной семьи";
        worksheet.Cells[2, 9].Value = " ";
        worksheet.Cells[1, 10].Value = "из неполной семьи (один из родителей умер, родители разведены)";
        worksheet.Cells[2, 10].Value = " ";
        worksheet.Cells[1, 11, 1, 14].Value = "сироты";
        worksheet.Cells[1, 11, 1, 14].Style.Font.Bold = true;
        worksheet.Cells[2, 11].Value = "всего, из них";
        worksheet.Cells[2, 11].Style.Font.Bold = true;
        worksheet.Cells[2, 12].Value = "на гос. обеспечении в БГУ";
        worksheet.Cells[2, 13].Value = "утратили последнего (единственного) родителя во время учебы в БГУ";
        worksheet.Cells[2, 14].Value = "приравненные к категории детей сирот";
        worksheet.Cells[1, 15, 1, 19].Value = "инвалиды";
        worksheet.Cells[1, 15, 1, 19].Style.Font.Bold = true;
        worksheet.Cells[2, 15].Value = "всего, из них";
        worksheet.Cells[2, 15].Style.Font.Bold = true;
        worksheet.Cells[2, 16].Value = "дети-инвалиды";
        worksheet.Cells[2, 17].Value = "инвалиды I группы";
        worksheet.Cells[2, 18].Value = "инвалиды II группы";
        worksheet.Cells[2, 19].Value = "инвалиды III группы";
        worksheet.Cells[1, 20, 1, 26].Value = "семейное положение";
        worksheet.Cells[1, 20, 1, 26].Style.Font.Bold = true;
        worksheet.Cells[2, 20].Value = "студентов, состоящих в браке, из них:";
        worksheet.Cells[2, 20].Style.Font.Bold = true;
        worksheet.Cells[2, 21].Value = "супруг(а)-студент(ка) БГУ, имеют детей";
        worksheet.Cells[2, 22].Value = "супруг(а)-студент(ка) БГУ, нет детей";
        worksheet.Cells[2, 23].Value = "супруг(а)-студент(ка) другого ВУЗа, имеют детей";
        worksheet.Cells[2, 24].Value = "супруг(а)-студент(ка) другого ВУЗа, нет детей";
        worksheet.Cells[2, 25].Value = "супруг(а) не студент, имеют детей";
        worksheet.Cells[2, 26].Value = "супруг(а) не студент, нет детей";
        worksheet.Cells[1, 27].Value = "имеет детей, не состоя в браке";
        worksheet.Cells[2, 27].Value = " ";
        worksheet.Cells[1, 28, 1, 33].Value = "проживают в общежитии";
        worksheet.Cells[1, 28, 1, 33].Style.Font.Bold = true;
        worksheet.Cells[2, 28].Value = "всего, из них:";
        worksheet.Cells[2, 28].Style.Font.Bold = true;
        worksheet.Cells[2, 29].Value = "обучаются платно";
        worksheet.Cells[2, 30].Value = "обучаются на бюджете";
        worksheet.Cells[2, 31].Value =
            "имеющие льготы в соответствии с Законом «О социальной защите граждан, пострадавших от катастрофы на Чернобыльской АЭС»";
        worksheet.Cells[2, 32].Value = "из многодетной семьи";
        worksheet.Cells[2, 33].Value = "инвалиды";
        worksheet.Cells[1, 34].Value = "достижения в спорте, науке";
        worksheet.Cells[1, 35].Value = "член БРСМ";
        worksheet.Cells[1, 36].Value = "член профсоюза";
        worksheet.Cells[1, 37].Value = "участие в органах студенческого самоуправления";

        worksheet.Cells[1, 3, 2, 37].Style.WrapText = true;
        worksheet.Cells[1, 3, 2, 37].Style.TextRotation = 90;
        worksheet.Cells[1, 3, 2, 37].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
        worksheet.Cells[1, 3, 2, 37].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

        worksheet.Cells[1, 11, 1, 26].Style.TextRotation = 0;
        worksheet.Cells[1, 28, 1, 33].Style.TextRotation = 0;
        worksheet.Cells[1, 1, 2, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
    }

    public static void FormatColumns(ExcelWorksheet worksheet)
    {
        for (var row = 2; !string.IsNullOrEmpty(worksheet.Cells[row, 1].Text); row++)
        {
            var formattedValueE = FormatCellValue(worksheet.Cells[row, 5].Text);
            worksheet.Cells[row, 5].Value = formattedValueE;
        }
    }

    public static void FillColumns(ExcelWorksheet groupWorksheet, ExcelWorksheet mainInformationPage, int rowCounter,
        int row)
    {
        groupWorksheet.Cells[rowCounter, 1].Value = rowCounter - 2;
        groupWorksheet.Cells[rowCounter, 2].Value = mainInformationPage.Cells[row, 2].Text;


        if (mainInformationPage.Cells[row, 9].Text == "Минск")
        {
            groupWorksheet.Cells[rowCounter, 3].Value = 1;
        }
        else
        {
            if (mainInformationPage.Cells[row, 7].Text == "Республика Беларусь")
                groupWorksheet.Cells[rowCounter, 4].Value = 1;
            else
                groupWorksheet.Cells[rowCounter, 6].Value = 1;
        }

        if (mainInformationPage.Cells[row, 11].Text == "Да (см. следующий вопрос)")
            groupWorksheet.Cells[rowCounter, 5].Value = 1;
        if (mainInformationPage.Cells[row, 6].Text == "Платная")
            groupWorksheet.Cells[rowCounter, 7].Value = 1;
        else
            groupWorksheet.Cells[rowCounter, 8].Value = 1;
        if (mainInformationPage.Cells[row, 14].Text == "Многодетная") groupWorksheet.Cells[rowCounter, 9].Value = 1;
        if (mainInformationPage.Cells[row, 13].Text != "Полная семья") groupWorksheet.Cells[rowCounter, 10].Value = 1;
        if (mainInformationPage.Cells[row, 15].Text != "Нет")
        {
            groupWorksheet.Cells[rowCounter, 11].Value = 1;

            if (mainInformationPage.Cells[row, 15].Text == "Да, на гос. обеспечении в БГУ")
                groupWorksheet.Cells[rowCounter, 12].Value = 1;
            if (mainInformationPage.Cells[row, 15].Text ==
                "Да, утратил последнего (единственного) родителя во время учёбы в БГУ")
                groupWorksheet.Cells[rowCounter, 13].Value = 1;
            if (mainInformationPage.Cells[row, 15].Text == "Да, другая категория")
                groupWorksheet.Cells[rowCounter, 14].Value = 1;
        }

        if (mainInformationPage.Cells[row, 16].Text != "Нет")
        {
            groupWorksheet.Cells[rowCounter, 15].Value = 1;

            if (mainInformationPage.Cells[row, 16].Text == "Да, ребёнок-инвалид")
                groupWorksheet.Cells[rowCounter, 16].Value = 1;
            if (mainInformationPage.Cells[row, 16].Text == "Да, инвалид I группы")
                groupWorksheet.Cells[rowCounter, 17].Value = 1;
            if (mainInformationPage.Cells[row, 16].Text == "Да, инвалид II группы")
                groupWorksheet.Cells[rowCounter, 18].Value = 1;
            if (mainInformationPage.Cells[row, 16].Text == "Да, инвалид III группы")
                groupWorksheet.Cells[rowCounter, 19].Value = 1;
        }

        if (mainInformationPage.Cells[row, 17].Text != "Не в браке")
        {
            groupWorksheet.Cells[rowCounter, 20].Value = 1;

            if (mainInformationPage.Cells[row, 15].Text == "В браке со студентом (студенткой) БГУ, есть дети")
                groupWorksheet.Cells[rowCounter, 21].Value = 1;
            if (mainInformationPage.Cells[row, 15].Text == "В браке со студентом (студенткой) БГУ, нет детей")
                groupWorksheet.Cells[rowCounter, 22].Value = 1;
            if (mainInformationPage.Cells[row, 15].Text == "В браке со студентом (студенткой) другого ВУЗа, есть дети")
                groupWorksheet.Cells[rowCounter, 23].Value = 1;
            if (mainInformationPage.Cells[row, 15].Text == "В браке со студентом (студенткой) другого ВУЗа, нет детей")
                groupWorksheet.Cells[rowCounter, 24].Value = 1;
            if (mainInformationPage.Cells[row, 15].Text == "В браке с не студентом (студенткой), есть дети")
                groupWorksheet.Cells[rowCounter, 25].Value = 1;
            if (mainInformationPage.Cells[row, 15].Text == "В браке с не студентом (студенткой), нет детей")
                groupWorksheet.Cells[rowCounter, 26].Value = 1;
        }

        if (mainInformationPage.Cells[row, 10].Text == "Проживаю")
        {
            groupWorksheet.Cells[rowCounter, 28].Value = 1;

            if (mainInformationPage.Cells[row, 6].Text == "Платная")
                groupWorksheet.Cells[rowCounter, 29].Value = 1;
            else
                groupWorksheet.Cells[rowCounter, 30].Value = 1;
            if (mainInformationPage.Cells[row, 11].Text == "Да (см. следующий вопрос)")
                groupWorksheet.Cells[rowCounter, 31].Value = 1;
            if (mainInformationPage.Cells[row, 14].Text == "Многодетная")
                groupWorksheet.Cells[rowCounter, 32].Value = 1;
            if (mainInformationPage.Cells[row, 16].Text != "Нет") groupWorksheet.Cells[rowCounter, 33].Value = 1;
        }

        if (mainInformationPage.Cells[row, 18].Text == "Есть") groupWorksheet.Cells[rowCounter, 34].Value = 1;
        if (mainInformationPage.Cells[row, 19].Text == "Да") groupWorksheet.Cells[rowCounter, 35].Value = 1;
        if (mainInformationPage.Cells[row, 20].Text == "Да") groupWorksheet.Cells[rowCounter, 36].Value = 1;
        if (mainInformationPage.Cells[row, 21].Text == "Да") groupWorksheet.Cells[rowCounter, 37].Value = 1;
    }

    public static void CreateTableBorders(ExcelWorksheet worksheet)
    {
        using (var headerCells = worksheet.Cells[3, 1, worksheet.Dimension.End.Row, worksheet.Dimension.End.Column])
        {
            headerCells.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            headerCells.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            headerCells.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            headerCells.Style.Border.Left.Style = ExcelBorderStyle.Thin;
        }
    }

    public static void FillingCommonInformation(ExcelWorksheet worksheet, ExcelWorksheet commonPage)
    {
        for (var col = 3; col <= worksheet.Dimension.End.Column; col++)
        {
            var columnCounter = 0;
            for (var row = 2; row <= worksheet.Dimension.End.Row; row++)
            {
                var cellValue = worksheet.Cells[row, col].Text;

                if (cellValue == "1") columnCounter++;
            }

            if (col == 3)
            {
                worksheet.Cells[worksheet.Dimension.End.Row + 1, col].Value = columnCounter;
                worksheet.Cells[worksheet.Dimension.End.Row, col - 1].Value = "ВСЕГО";
                worksheet.Cells[worksheet.Dimension.End.Row, col - 1].Style.Font.Bold = true;
                commonPage.Cells[commonPage.Dimension.End.Row + 1, col].Value = columnCounter;
            }
            else
            {
                worksheet.Cells[worksheet.Dimension.End.Row, col].Value = columnCounter;
                commonPage.Cells[commonPage.Dimension.End.Row, col].Value = columnCounter;
            }
        }

        commonPage.Cells[commonPage.Dimension.End.Row, 2].Value = worksheet.Name;
    }

    private static string FormatCellValue(string value)
    {
        var chars = value.ToCharArray();
        var formattedValue = "";

        for (var i = 0; i < chars.Length; i++)
            if (char.IsLetter(chars[i]))
                formattedValue += char.ToUpper(chars[i]);
            else if (char.IsDigit(chars[i]))
                formattedValue += chars[i];
            else
                formattedValue += " " + chars[i] + " ";

        formattedValue = formattedValue.Replace("  ", " ");

        return formattedValue.Trim();
    }
}