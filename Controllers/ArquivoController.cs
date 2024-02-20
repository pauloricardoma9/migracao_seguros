using Microsoft.AspNetCore.Mvc;
using migracao_seguros.Services;

namespace migracao_seguros.Controllers;

[ApiController]
[Route("arquivo")]
public class ArquivoController : ControllerBase, IArquivoController
{
    private readonly IArquivoService arquivoService;

    public ArquivoController(IArquivoService arquivoService)
    {
        this.arquivoService = arquivoService;
    }

    [HttpPost("processar")]
    public IActionResult Processar(IFormFile arquivoXlsx)
    {
        var (arquivo, erro) = arquivoService.Processar(arquivoXlsx);

        if(arquivo != null)
        {
            return arquivo;
        }

        return BadRequest(erro);   
    }
}
