using ExcelImporter.Services;
using Microsoft.AspNetCore.Mvc;

namespace ExcelImporter.Web.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ImportController : ControllerBase
    {
        protected readonly ILogger<ImportController> _logger;
        private readonly IImportService _importService;

        public ImportController(IImportService importService, ILogger<ImportController> logger = null!)
        {
            _logger = logger;
            _importService = importService;
        }

        /// <summary>
        /// Test posting data. 
        /// </summary>
        /// <param name="data">Any object convertible to JSON.</param>
        /// <returns>Message for success and echoed data.</returns>
        [HttpPost("test")]
        [ProducesResponseType(200)]
        [ProducesResponseType(400)]
        [ProducesResponseType(500)]
        public IActionResult Test([FromBody] object data)
        {
            try
            {
                var json = Newtonsoft.Json.JsonConvert.SerializeObject(data);
                string result = $"Test succeeded! data: {json}";
                
                _logger.LogInformation(result);
                return Ok(result);
            }
            catch(Exception ex)
            {
                _logger.LogError(ex, "Error on import test");
                return StatusCode(StatusCodes.Status500InternalServerError, "Error on import test");
            }
        }


        /// <summary>
        /// Test posting file. 
        /// </summary>
        /// <returns>Message for success and echoed data.</returns>
        [HttpPost("testFile")]
        [ProducesResponseType(200)]
        [ProducesResponseType(400)]
        [ProducesResponseType(500)]
        public async Task<IActionResult> TestFile()
        {
            try
            {
                if (Request.Form.Files == null || Request.Form.Files.Count() == 0 || Request.Form.Files[0] == null)
                {
                    return BadRequest("Missing import file!");
                }

                var file = await _importService.ParseFile(Request.Form.Files[0]);

                var json = Newtonsoft.Json.JsonConvert.SerializeObject(file);
                string result = $"Test succeeded! file: {json}";

                return Ok(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error on importing a test file");
                return StatusCode(StatusCodes.Status500InternalServerError, "Error on importing a test file");
            }
        }


    }
}
