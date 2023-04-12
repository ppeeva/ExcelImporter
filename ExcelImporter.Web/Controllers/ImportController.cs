using Microsoft.AspNetCore.Mvc;

namespace ExcelImporter.Web.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ImportController : ControllerBase
    {
        public ImportController()
        {
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
                if (data == null)
                {
                    return BadRequest("Missing data!");
                }

                var json = Newtonsoft.Json.JsonConvert.SerializeObject(data);
                string result = $"Test succeeded! data: {json}";

                return Ok(result);
            }
            catch 
            {
                return StatusCode(StatusCodes.Status500InternalServerError, "Error on import test");
            }
        }


    }
}
