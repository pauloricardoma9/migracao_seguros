using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using migracao_seguros.Services;

namespace teste.Pages;

public class IndexModel : PageModel
{
    private readonly IArquivoService arquivoService;

    public IndexModel(IArquivoService arquivoService)
    {
        this.arquivoService = arquivoService;
    }

    public IActionResult OnPost(IFormFile? arquivoRecebido)
    {
        var (arquivo, erro) = arquivoService.Processar(arquivoRecebido!);

        if(arquivo != null)
        {
            return arquivo;
        }

        ModelState.AddModelError(string.Empty, erro ?? "Ocorreu um erro ao processar o arquivo.");
        return Page();
    }
}
