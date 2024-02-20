using Microsoft.AspNetCore.Mvc;

namespace migracao_seguros.Controllers;
public interface IArquivoController
{
    public IActionResult Processar(IFormFile arquivoXlsx);
}