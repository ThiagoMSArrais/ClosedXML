using ClosedXML.Excel;

namespace Excel
{
    public class ExportacaoExcel
    {
        public static void Extracao()
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Sample Sheet");
                worksheet.Cell("A1").Value = "Hello World!";
                worksheet.Cell("A2").FormulaA1 = "=MID(A1, 7, 5)";
                worksheet.Cell("A3").Value = "Bom dia";
                worksheet.Cell("B3").Value = "Boa tarde!";

                workbook.SaveAs("HelloWorld.xlsx"); 
            }
        }
    }
}
