using Microsoft.AspNetCore.Mvc;

namespace migracao_seguros.Services;
public interface IArquivoService
{
    public (FileContentResult? arquivoConvertido, string? erro) Processar(IFormFile arquivoRecebido);
}