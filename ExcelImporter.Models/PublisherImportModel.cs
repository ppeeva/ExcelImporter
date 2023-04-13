using ExcelImporter.Shared.Attributes;
using System.ComponentModel.DataAnnotations;

namespace ExcelImporter.Models
{
    public class PublisherImportModel
    {
        [ExportTemplate("Код на издател")]
        [Required]
        [Range(1, int.MaxValue)]
        public int Id { get; set; }

        [ExportTemplate("Име на издател")]
        [Required(AllowEmptyStrings = false)]
        public string? Name { get; set; }

        [ExportTemplate("Местоположение")]
        public string? Location { get; set; }

        
        public IEnumerable<BookImportModel>? Books { get; set; }
    }
}
