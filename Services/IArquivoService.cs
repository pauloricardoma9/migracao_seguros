using Microsoft.AspNetCore.Mvc;

namespace migracao_seguros.Services;
public interface IArquivoService
{
    public Task<(FileContentResult? arquivoConvertido, string? erro, List<int>? linhasNaoProcessadas)> Processar(IFormFile arquivoRecebido);
}