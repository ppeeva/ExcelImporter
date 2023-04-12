using ExcelImporter.Models;
using Microsoft.AspNetCore.Http;

namespace ExcelImporter.Services
{
    public interface IImportService
    {
        Task<FileModel> ParseFile(IFormFile file);
    }
}
