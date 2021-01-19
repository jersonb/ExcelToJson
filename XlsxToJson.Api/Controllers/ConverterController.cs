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
        [Route("object")]
        public IActionResult Post([FromBody] object json)
        {
            JsonToXlsx jsonToXlsx = json.ToString();

            return File(jsonToXlsx.MemoryStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "test.xlsx");
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

            XlsxToJson xlsxToJson = stream;

            return Ok(xlsxToJson.Json);
        }
    }
}