using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Mvc;
using webContratoQR.Models;

namespace webContratoQR.Controllers
{
    public class HomeController() : Controller
    {
        private readonly IConfiguration _configuration = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build();

        public ActionResult Index()
        {
            FileExcelModel fileExcelModel = new FileExcelModel();
            ViewData["qr"] = "";
            ViewBag.QR = "";
            fileExcelModel.CodigoQR = "";
            fileExcelModel.NombreFuncionario = "";
            fileExcelModel.UrlContrato = "";

            return View(fileExcelModel);
        }


        [HttpPost]
        public Task<IActionResult> Index(string texto)
        {
            List<FileExcel> funcionarios = new List<FileExcel>();
            string RutFuncionario = "";
            string NombreFuncionario = "";
            string UrlContrato = "";
            string? filePath = ""; // _configuration.GetValue<string>("NombreExcel").ToString();
            FileExcelModel fileExcelModel = new FileExcelModel();

            if (Directory.GetFiles(Path.Combine("wwwroot", "Uploads")).Length > 0)
            {
                filePath = Directory.GetFiles(Path.Combine("wwwroot", "Uploads"))[0];
            }
            else
            {
                fileExcelModel.Mensaje = "No se ha subido el archivo Excel.";
                fileExcelModel.IsError = "SI";
                return Task.FromResult<IActionResult>(View());
            }

            if (texto == null)
            {
                ViewData["mensaje"] = "Introduzca Rut del Funcionario";
                return Task.FromResult<IActionResult>(View());
            }

            ViewData["qr"] = "";


            // 1. Abre el documento de Excel en modo de solo lectura
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
            {
                //create the object for workbook part  
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                Sheets thesheetcollection = workbookPart.Workbook.GetFirstChild<Sheets>();

                //using for each loop to get the sheet from the sheetcollection  
                foreach (Sheet thesheet in thesheetcollection)
                {
                    //statement to get the worksheet object by using the sheet id  
                    Worksheet theWorksheet = ((WorksheetPart)workbookPart.GetPartById(thesheet.Id)).Worksheet;

                    SheetData thesheetdata = theWorksheet.GetFirstChild<SheetData>();
                    foreach (Row thecurrentrow in thesheetdata)
                    {
                        foreach (Cell thecurrentcell in thecurrentrow)
                        {
                            if (thecurrentcell.CellReference != "A1" && thecurrentcell.CellReference != "B1" && thecurrentcell.CellReference != "C1")
                            {
                                //statement to take the integer value  
                                string currentcellvalue = string.Empty;
                                if (thecurrentcell.DataType != null)
                                {
                                    if (thecurrentcell.DataType == CellValues.SharedString)
                                    {
                                        int id;

                                        if (Int32.TryParse(thecurrentcell.InnerText, out id))
                                        {
                                            SharedStringItem item = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
                                            if (item.Text != null && thecurrentcell.CellReference.ToString().Substring(0, 1) == "A")
                                            {
                                                RutFuncionario = item.Text.Text.Replace("-", "");
                                            }
                                            else if (item.Text != null && thecurrentcell.CellReference.ToString().Substring(0, 1) == "B")
                                            {
                                                NombreFuncionario = item.Text.Text;
                                            }
                                            else if (item.Text != null && thecurrentcell.CellReference.ToString().Substring(0, 1) == "C")
                                            {
                                                UrlContrato = item.Text.Text;
                                            }

                                        }
                                    }
                                }
                            }

                        }

                        funcionarios.Add(new FileExcel { RutFuncionario = RutFuncionario, NombreFuncionario = NombreFuncionario, UrlContrato = UrlContrato });
                    }
                }
            }

            fileExcelModel.Funcionarios = funcionarios;
            fileExcelModel.Funcionario = fileExcelModel.Funcionarios.Where(f => f.RutFuncionario == texto.Replace("-", "").Replace(".", "")).FirstOrDefault();

            Helpers.HelperQR helperQR = new Helpers.HelperQR();
            //ViewData["qr"] = helperQR.GenerateQRCode(fileExcelModel.Funcionario != null ? fileExcelModel.Funcionario.UrlContrato : "No encontrado");
            fileExcelModel.CodigoQR = helperQR.GenerateQRCode(fileExcelModel.Funcionario != null ? fileExcelModel.Funcionario.UrlContrato : "No encontrado");
            fileExcelModel.NombreFuncionario = fileExcelModel.Funcionario == null ? "No encontrado" : fileExcelModel.Funcionario.NombreFuncionario;
            fileExcelModel.UrlContrato = fileExcelModel.Funcionario == null ? "No encontrado" : fileExcelModel.Funcionario.UrlContrato;

            return Task.FromResult<IActionResult>(View(fileExcelModel));
        }
    }
}

