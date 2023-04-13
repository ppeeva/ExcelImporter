using ExcelImporter.Shared.Attributes;
using System.ComponentModel.DataAnnotations;

namespace ExcelImporter.Models
{
    public class BookImportModel
    {
        [ExportTemplate("Код на издател")]
        [Required]
        [Range(1, int.MaxValue)]
        public int PublisherId { get; set; }

        [ExportTemplate("ISBN")]
        [Required(AllowEmptyStrings = false)]
        public string? ISBN { get; set; }

        [ExportTemplate("Заглавие")]
        [Required(AllowEmptyStrings = false)]
        public string? Title { get; set; }



        [ExportTemplate("Дата на регистриране - ден")]
        [Range(1, 31)]
        public int? RegisterDateDay { get; set; }

        [ExportTemplate("Дата на регистриране - месец")]
        [Range(1, 12)]
        public int? RegisterDateMonth { get; set; }

        [ExportTemplate("Дата на регистриране - година")]
        [Required]
        [Range(1, 3000)]
        public int RegisterDateYear { get; set; }



        [ExportTemplate("Езици", Shared.ExcelDropDownListType.Language, true, "LanguageCodes")]
        public string? Languages { get; set; }

        [ExportTemplate("Описание")]
        public string? Description { get; set; }



        public IEnumerable<string>? LanguageCodes { get; set; }

        public bool IsRegisterDateValid
        {
            get
            {
                try
                {
                    var date = new DateTime(RegisterDateYear, RegisterDateMonth ?? 1, RegisterDateDay ?? 1);
                    return true;
                }
                catch
                {
                    return false;
                }
            }
        }
    }
}
