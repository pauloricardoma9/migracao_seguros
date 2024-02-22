using System.Diagnostics;
using Microsoft.AspNetCore.Mvc;
using migracao_seguros.Models;
using migracao_seguros.Services;

namespace migracao_seguros.Controllers;

public class ArquivoController : Controller
{
    private readonly ILogger<ArquivoController> _logger;
    private readonly IArquivoService arquivoService;

    public ArquivoController(ILogger<ArquivoController> logger, IArquivoService arquivoService)
    {
        _logger = logger;
        this.arquivoService = arquivoService;
    }

    public IActionResult Index()
    {
        return View();
    }

    public async Task<IActionResult> Converter(IFormFile arquivoOrigem)
    {
        var (arquivo, erro, linhasNaoProcessadas) = await arquivoService.Processar(arquivoOrigem);
        
        var model = new {
            Erro = erro,
            Arquivo = arquivo,
            LinhasNaoProcessadas = linhasNaoProcessadas == null ? "" : string.Join(", ", linhasNaoProcessadas)
        };

        if(arquivo != null)
        {
            return arquivo;
        }

        return View("Index", model);   
    }

    // public IActionResult Download()
    // {
    //     return View();
    // }

    [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
    public IActionResult Error()
    {
        return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
    }
}
