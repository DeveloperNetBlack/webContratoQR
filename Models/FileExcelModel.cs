namespace webContratoQR.Models
{
    public class FileExcelModel
    {
        public List<FileExcel> Funcionarios { get; set; } = new List<FileExcel>();
        public FileExcel? Funcionario { get; set; }
        public string? CodigoQR { get; set; }
        public string? NombreFuncionario { get; set; }
        public string? UrlContrato { get; set; }
        public string? Mensaje { get; set; }
        public string? IsError { get; set; }
    }
}
