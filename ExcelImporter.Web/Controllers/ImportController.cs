using Microsoft.AspNetCore.Mvc;

namespace ExcelImporter.Web.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ImportController : ControllerBase
    {
        protected readonly ILogger<ImportController> _logger;

        public ImportController(ILogger<ImportController> logger = null!)
        {
            _logger = logger;
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


    }
}
