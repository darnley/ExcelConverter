using System.IO;
using Converter.ExcelConverter.Services;
using Microsoft.AspNetCore.Mvc;

namespace Converter.Presentation.Web.Api.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ConverterController : ControllerBase
    {
        // GET api/converter?path={string}
        [HttpGet]
        public ActionResult Get([FromQuery] string path, [FromServices] IExcelConverterService service)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            using (var stream = System.IO.File.Open(path, FileMode.Open, FileAccess.Read))
            {
                return Ok(service.ConvertExcelToDictionary(stream));
            }
        }
    }
}
