using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;

namespace webContratoQR.Helpers
{
    public class ReadFileExcel
    {
        public async Task<List<string>> ReadExcelFileAsync()
        {

            List<string> data = new List<string>();


            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open("D:\\Proyectos\\QUILICURA\\FUNCIONARIOS SALUD 2025.xlsx", false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                IEnumerable<Sheet> sheets = workbookPart.Workbook.Descendants<Sheet>();
                string relationshipId = sheets.First().Id.Value;
                WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(relationshipId);
                Worksheet worksheet = worksheetPart.Worksheet;

                // Puedes iterar por las celdas usando LINQ
                foreach (Row row in worksheet.Descendants<Row>())
                {
                    foreach (Cell cell in row.Descendants<Cell>())
                    {
                        // Leer el valor de la celda
                        string cellValue = GetCellValue(cell, workbookPart);
                        if (!string.IsNullOrEmpty(cellValue) && cellValue.ToUpper() != "RUT" && cell.CellReference.Value.Substring(0, 1) == "A")
                        {
                            data.Add(cellValue);
                        }
                    }
                }
            }

            return await Task.FromResult(data);
        }

        // Método para obtener el valor de una celda, manejando el caso de SharedString
        private static string GetCellValue(Cell cell, WorkbookPart workbookPart)
        {
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                // Obtener el valor de SharedStrings
                var sharedStringTable = workbookPart.SharedStringTablePart.SharedStringTable;
                int ssid = int.Parse(cell.CellValue.InnerXml);
                return sharedStringTable.Elements<SharedStringItem>().ElementAt(ssid).InnerText;
            }
            return cell.CellValue != null ? cell.CellValue.InnerXml : string.Empty;
        }

    }

}
