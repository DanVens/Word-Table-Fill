using System;
using System.IO;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;
using WebApplication1.Command;

namespace WebApplication1.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class DocumentController : ControllerBase
    {
        private readonly string _templatePath;
        private readonly PostDocumentCommand _command = new();

        public DocumentController(IConfiguration config, IHostEnvironment env)
        {
            // Read from appsettings:  "Templates:Word": "Templates/mok_NMPP_protokolas.docx"
            var configured = config["Templates:Word"];
            if (string.IsNullOrWhiteSpace(configured))
                throw new InvalidOperationException("Templates:Word is not configured.");

            // Resolve relative path against content root
            _templatePath = Path.IsPathRooted(configured)
                ? configured
                : Path.Combine(env.ContentRootPath, configured);
        }

        [HttpPost("fill")]
        public IActionResult Fill([FromBody] FillDocumentDto dto)
        {
            if (!System.IO.File.Exists(_templatePath))
                return NotFound($"Template not found at '{_templatePath}'.");

            using var fs = System.IO.File.OpenRead(_templatePath);
            var result = _command.Execute(fs, dto);
            result.Position = 0;

            return File(
                result,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "FilledDocument.docx");
        }
    }
}