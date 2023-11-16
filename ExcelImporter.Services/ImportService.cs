using ExcelImporter.Models;
using ExcelImporter.Shared;
using ExcelImporter.Shared.Excel;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using OfficeOpenXml;
using System.Text;

namespace ExcelImporter.Services
{
    public class ImportService : IImportService
    {
        protected readonly ILogger _logger;
        private readonly ImportSettings _importSettings;
        private readonly IModelValidationService _modelValidationService;

        public ImportService(IOptions<ImportSettings> importSettings, 
            ILogger<ImportService> logger, 
            IModelValidationService modelValidationService)
        {
            _logger = logger;
            _importSettings = importSettings.Value;
            _modelValidationService = modelValidationService;
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


        private IList<ExcelDDL> GetDropDownListsCollection(ExcelDropDownListType[] ddlTypes)
        {
            List<ExcelDDL> lists = new List<ExcelDDL>();

            foreach (ExcelDropDownListType ddlType in ddlTypes)
            {
                switch (ddlType)
                {
                    case ExcelDropDownListType.Language:
                        List<ExcelDDLItem> languages = new List<ExcelDDLItem>{
                            new ExcelDDLItem { Value = "en", DisplayName = "Английски" },
                            new ExcelDDLItem { Value = "bg", DisplayName = "Български" },
                            new ExcelDDLItem { Value = "de", DisplayName = "Немски" },
                            new ExcelDDLItem { Value = "ru", DisplayName = "Руски" }
                        };

                        lists.Add(new ExcelDDL { ListType = ExcelDropDownListType.Language, Collection = languages });
                        break;

                    default:
                        break;
                }
            }

            return lists;
        }



        public async Task<Tuple<bool, object>> ImportBooksFromExcelAsync(IFormFile model)
        {
            try
            {
                _logger.LogInformation($"Start importing file {model.FileName}");

                string publisherSheetName = _importSettings.PublishersSheetName ?? "";
                string booksSheetName = _importSettings.BooksSheetName ?? "";

                string fileName = model.FileName;
                var xlsx = model;
                string error = "";
                StringBuilder errorsList = new StringBuilder();

                List<PublisherImportModel> publisherImportModels = new List<PublisherImportModel>();
                List<BookImportModel> bookImportModels = new List<BookImportModel>();

                IList<ExcelDDL> ddlLists = GetDropDownListsCollection(
                    new ExcelDropDownListType[] {
                        ExcelDropDownListType.Language
                    });


                if (xlsx != null && xlsx.Length > 0)
                {
                    using (var stream = new MemoryStream())
                    {
                        await xlsx.CopyToAsync(stream);
                        using (var package = new ExcelPackage(stream))
                        {
                            ExcelWorksheet publisherWorksheet = package.Workbook.Worksheets[publisherSheetName];
                            if (publisherWorksheet == null)
                            {
                                error = $"Sheet {publisherSheetName} is missing!";
                                _logger.LogError(error);
                                return new Tuple<bool, object>(false, error);
                            }

                            ImportPublishers(publisherWorksheet, ddlLists, error, errorsList, publisherImportModels);
                            if (errorsList.Length > 0)
                            {
                                _logger.LogError(errorsList.ToString());
                                return new Tuple<bool, object>(false, errorsList.ToString());
                            }

                            if (publisherImportModels.Count == 0)
                            {
                                error = "No archival entities data found!";
                                _logger.LogError(error);
                                return new Tuple<bool, object>(false, error);
                            }



                            ExcelWorksheet bookWorksheet = package.Workbook.Worksheets[booksSheetName];
                            if (bookWorksheet == null)
                            {
                                error = $"Sheet {booksSheetName} is missing!";
                                _logger.LogError(error);
                                return new Tuple<bool, object>(false, error);
                            }

                            ImportBooks(bookWorksheet, ddlLists, error, errorsList, bookImportModels, publisherImportModels);
                        }

                        if (errorsList.Length > 0)
                        {
                            _logger.LogError(errorsList.ToString());
                            return new Tuple<bool, object>(false, errorsList.ToString());
                        }

                        _logger.LogInformation($"Import of file {model.FileName} succeeded");
                        return new Tuple<bool, object>(true, publisherImportModels);
                    }
                }
                else
                {
                    error = "ImportEmptyFile";
                    _logger.LogError(error);
                    return new Tuple<bool, object>(false, error);
                }
            }
            catch
            {
                throw;
            }
        }


        private void ImportPublishers(ExcelWorksheet publisherWorksheet, IList<ExcelDDL> ddlLists, string error, StringBuilder errorsList, List<PublisherImportModel> publisherImportModels)
        {
            int publisherRowCnt = publisherWorksheet.Dimension != null ? publisherWorksheet.Dimension.Rows : 0;
            int publisherColCnt = publisherWorksheet.Dimension != null ? publisherWorksheet.Dimension.Columns : 0;

            if (publisherRowCnt > 2 && publisherColCnt > 0)
            {
                for (int row = 3; row <= publisherRowCnt; row++)
                {
                    PublisherImportModel publisherImportModel = ExcelHelper.GetModelFromRow<PublisherImportModel>(publisherWorksheet, 1, publisherColCnt, row, 1, ddlLists);
                    if (publisherImportModel == null)
                        break;

                    string validationErrors = _modelValidationService.ValidateToString(publisherImportModel);
                    if (String.IsNullOrWhiteSpace(validationErrors))
                    {
                        bool hasError = false;
                        string validateDDLError;
                        if (!ExportModelHelper.AreDropDownValuesValid(publisherImportModel, ddlLists, out validateDDLError))
                        {
                            error = $"Invalid nomenclature data {validateDDLError} for publisher on row {row}";
                            errorsList.AppendLine(error);
                            hasError = true;
                        }

                        if (publisherImportModels.Where(x => x.Id == publisherImportModel.Id).Any())
                        {
                            error = $"Duplicated publisher number {publisherImportModel.Id} on row {row}!";
                            errorsList.AppendLine(error);
                            hasError = true;
                        }

                        if (!hasError)
                        {
                            publisherImportModels.Add(publisherImportModel);
                        }
                    }
                    else
                    {
                        error = $"Invalid publisher data on row {row}: {validationErrors}";
                        errorsList.AppendLine(error);
                    }
                }
            }
        }

        private void ImportBooks(ExcelWorksheet bookWorksheet, IList<ExcelDDL> ddlLists, string error, StringBuilder errorsList, List<BookImportModel> bookImportModels, List<PublisherImportModel> publisherImportModels)
        {
            int bookRowCnt = bookWorksheet.Dimension != null ? bookWorksheet.Dimension.Rows : 0;
            int bookColCnt = bookWorksheet.Dimension != null ? bookWorksheet.Dimension.Columns : 0;

            if (bookRowCnt > 2 && bookColCnt > 0)
            {
                for (int row = 3; row <= bookRowCnt; row++)
                {
                    BookImportModel bookImportModel = ExcelHelper.GetModelFromRow<BookImportModel>(bookWorksheet, 1, bookColCnt, row, 1, ddlLists);
                    if (bookImportModel == null)
                        break;

                    string validationErrors = _modelValidationService.ValidateToString(bookImportModel);
                    if (String.IsNullOrWhiteSpace(validationErrors))
                    {
                        bool hasError = false;
                        string validateDDLError;
                        if (!ExportModelHelper.AreDropDownValuesValid(bookImportModel, ddlLists, out validateDDLError))
                        {
                            error = $"Invalid nomenclature data {validateDDLError} for book on row {row}";
                            errorsList.AppendLine(error);
                            hasError = true;
                        }

                        if (!bookImportModel.IsRegisterDateValid)
                        {
                            error = $"Invalid register date on row {row}";
                            errorsList.AppendLine(error);
                            hasError = true;
                        }

                        if (!hasError)
                        {
                            bookImportModels.Add(bookImportModel);

                            var publisher = publisherImportModels.FirstOrDefault(x => x.Id == bookImportModel.PublisherId);
                            if (publisher != null)
                            {
                                publisher.Books = publisher.Books ?? new List<BookImportModel>();
                                List<BookImportModel> booksList = publisher.Books.ToList();
                                booksList.Add(bookImportModel);
                                publisher.Books = booksList;
                            }
                            else
                            {
                                errorsList.AppendLine($"No matching publisher with Id {bookImportModel.PublisherId} found for book on row {row}.");
                            }
                        }
                    }
                    else
                    {
                        error = $"Invalid book data on row {row}: {validationErrors}";
                        errorsList.AppendLine(error);
                    }
                }
            }
        }


        public FileDownloadModel ExportFileTemplateToExcel()
        {
            IList<ExcelDDL> ddlLists = GetDropDownListsCollection(
                    new ExcelDropDownListType[] {
                        ExcelDropDownListType.Language,
                    });

            PublisherImportModel publisherModel = new PublisherImportModel();
            BookImportModel bookModel = new BookImportModel();

            string publisherSheetName = _importSettings.PublishersSheetName ?? "";
            string bookSheetName = _importSettings.BooksSheetName ?? "";


            var stream = new MemoryStream();
            using (var package = new ExcelPackage(stream))
            {
                var publisherWorkSheet = package.Workbook.Worksheets.Add(publisherSheetName);
                ExcelHelper.AddModelDataToSheet(package.Workbook, publisherWorkSheet, new List<PublisherImportModel>() { publisherModel }, ddlLists, false, true);

                var bookWorkSheet = package.Workbook.Worksheets.Add(bookSheetName);
                ExcelHelper.AddModelDataToSheet(package.Workbook, bookWorkSheet, new List<BookImportModel>() { bookModel }, ddlLists, false, true);

                package.Save();
            }
            stream.Position = 0;

            FileDownloadModel file = new FileDownloadModel()
            {
                Mimetype = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                Filename = "ImportFileTemplate.xlsx",
                Data = Convert.ToBase64String(stream.ToArray())
            };

            return file;
        }
    }
}
