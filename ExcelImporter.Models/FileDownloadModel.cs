
namespace ExcelImporter.Models
{
    public class FileDownloadModel
    {
        /// <summary>
        /// mimetype in the form 'major/minor'
        /// </summary>
        public string? Mimetype { get; set; }

        /// <summary>
        /// the name of the file to download
        /// </summary>
        public string? Filename { get; set; }

        /// <summary>
        /// the binary data as base64 to download
        /// </summary>
        public string? Data { get; set; }
    }
}
