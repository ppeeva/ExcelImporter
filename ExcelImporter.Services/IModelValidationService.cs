using System.ComponentModel.DataAnnotations;

namespace ExcelImporter.Services
{
    public interface IModelValidationService
    {
        List<ValidationResult> Validate<T>(T model) where T : class;

        string ValidateToString<T>(T model) where T : class;
    }
}
