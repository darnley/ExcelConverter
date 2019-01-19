using System.IO;
using Converter.ExcelConverter.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System.Threading.Tasks;

namespace Converter.Presentation.Web.Api.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ConverterController : ControllerBase
    {
        // GET api/converter?path={string}
        [HttpGet]
        public async Task<ActionResult> GetAsync(IFormFile file, [FromServices] IExcelConverterService service)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            if (file == null)
            {
                return BadRequest(new string[] { "There is no file in 'file' parameter in body." });
            }
            
            string temporaryFilePath = Path.GetTempFileName();

            if (file.Length > 0)
            {
                using (FileStream stream = new FileStream(temporaryFilePath, FileMode.Create))
                {
                    await file.CopyToAsync(stream);

                    return Ok(service.ConvertExcelToDictionary(stream));
                }
            }

            return BadRequest(new string[] { "The provided file is zero-byte length." });
        }
    }
}
