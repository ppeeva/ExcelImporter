using System.ComponentModel.DataAnnotations;
using System.Text;

namespace ExcelImporter.Services
{
    public class ModelValidationService : IModelValidationService
    {
        public List<ValidationResult> Validate<T>(T model) where T : class
        {
            var context = new ValidationContext(model, serviceProvider: null, items: null);
            var results = new List<ValidationResult>();

            Validator.TryValidateObject(model, context, results, true);

            return results;
        }

        public string ValidateToString<T>(T model) where T : class
        {
            var results = Validate(model);
            var msg = new StringBuilder();

            foreach (ValidationResult validationResult in results)
            {
                msg.AppendLine(validationResult.ErrorMessage ?? "");
            }

            return msg.ToString();
        }
    }
}
