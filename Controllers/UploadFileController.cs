using Microsoft.AspNetCore.Mvc;
using webContratoQR.Models;

namespace webContratoQR.Controllers
{
    public class UploadFileController : Controller
    {
        public IActionResult Index()
        {
            FileExcelModel fileExcelModel = new FileExcelModel();

            fileExcelModel.IsError = "";

            return View(fileExcelModel);
        }

        [HttpPost]
        public async Task<IActionResult> FileUpload(IFormFile file)
        {

            FileExcelModel fileExcelModel = new FileExcelModel();
            ViewBag.Inicio = "NO";

            try
            {
                if (file != null && file.Length > 0)
                {
                    var uploadsPath = Path.Combine("wwwroot", "Uploads"); // Asegúrate que esta carpeta existe
                    if (Directory.Exists(uploadsPath))
                    {
                        Directory.Delete(uploadsPath, true);
                        Directory.CreateDirectory(uploadsPath);
                    }
                    else
                    {
                        Directory.CreateDirectory(uploadsPath);
                    }

                    var filePath = Path.Combine(uploadsPath, file.FileName);

                    using (var stream = new FileStream(filePath, FileMode.Create))
                    {
                        await file.CopyToAsync(stream);
                    }

                    fileExcelModel.IsError = "NO";
                    fileExcelModel.Mensaje = "Archivo subido exitosamente.";
                }
            }
            catch (Exception ex)
            {
                fileExcelModel.Mensaje = ex.Message;
                fileExcelModel.IsError = "SI";
            }

            return View("Index", fileExcelModel);
        }
    }
}
