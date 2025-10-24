using ExcelDataReader;
using Microsoft.AspNetCore.Mvc;
using System.Data;
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

            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();

                    foreach(DataRow fila in result.Tables[0].Rows)
                    {
                        if (!fila[0].ToString().Equals("RUT", StringComparison.CurrentCultureIgnoreCase) && fila[1].ToString().ToUpper() != "NOMBRE" && fila[2].ToString().ToUpper() != "ENLACE")
                        {
                            funcionarios.Add(
                                new FileExcel 
                                { 
                                    RutFuncionario = fila[0].ToString(), 
                                    NombreFuncionario = fila[1].ToString(), 
                                    UrlContrato = fila[2].ToString() 
                                });
                        }
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

