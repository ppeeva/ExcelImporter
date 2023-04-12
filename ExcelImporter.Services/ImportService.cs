using ExcelImporter.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelImporter.Services
{
    public class ImportService : IImportService
    {
        protected readonly ILogger _logger;

        public ImportService(ILogger<ImportService> logger)
        {
            _logger = logger;
        }

        public async Task<FileModel> ParseFile(IFormFile file)
        {
            FileModel fileModel = await ParseAttachmentAsync(file);
            return fileModel;
        }

        private async Task<FileModel> ParseAttachmentAsync(IFormFile file)
        {
            var result = new FileModel();

            using (var stream = new MemoryStream())
            {
                await file.CopyToAsync(stream);
                result.Name = file.FileName;
                result.ContentType = file.ContentType;
                result.Type = file.FileName.Split('.').Last();
                result.Size = stream.Length;
                result.Content = stream.ToArray();
            }

            return result;
        }
    }
}
