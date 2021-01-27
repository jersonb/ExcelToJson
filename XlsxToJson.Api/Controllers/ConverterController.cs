using Converter;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace ExcelToJson.Api.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class ConverterController : ControllerBase
    {
        /// <summary>
        ///
        /// </summary>
        /// <param name="json"></param>
        /// <returns>Dowload file excel</returns>
        [HttpPost]
        [Route("object/{tyepe?}")]
        public IActionResult Post([FromBody] object json, string type = null)
        {
            var xlsx = ConvertTo.Xlsx(json.ToString());

            return File(xlsx.MemoryStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "test.xlsx");
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="upload"></param>
        /// <returns>Json</returns>
        [HttpPost]
        [Route("file")]
        public IActionResult Post(IFormFile upload)
        {
            var stream = upload.OpenReadStream();

            return Ok(ConvertTo.Json(stream));
        }
    }
}