using System.Text;
using Microsoft.AspNetCore.Mvc;
using Microsoft.VisualBasic.FileIO;
using migracao_seguros.Models;

namespace migracao_seguros.Services;

public class ArquivoService : IArquivoService
{
    private int sequencialRegistro = 1;
    private readonly string diretorioTemporario = Path.GetTempPath();
    private const char CARACTERE_COMPLEMENTAR = '0';
    private const char CARACTERE_COMPLEMENTAR_ALFANUMERICO = ' ';
    private const string CODIGO_BANCO = "047";
    private const string FORMATO_DATA_CONVERSAO = "yyyyMMdd";
    private const string TIPO_ARQUIVO_ENTRADA = ".csv";

    [HttpPost("processar")]
    public async Task<(FileContentResult? arquivoConvertido, string? erro, List<int>? linhasNaoProcessadas)> Processar(IFormFile arquivoEntrada)
    {
        var dataProcessamento = DateTime.Now.ToString(FORMATO_DATA_CONVERSAO);

        if(arquivoEntrada == null || arquivoEntrada.Length == 0)
        {
            return (null, $"Um arquivo {TIPO_ARQUIVO_ENTRADA} deve ser selecionado.", null);
        }

        var extensao = Path.GetExtension(arquivoEntrada.FileName);
        if(extensao != TIPO_ARQUIVO_ENTRADA)
        {
            return (null, $"O arquivo enviado deve ser do tipo {TIPO_ARQUIVO_ENTRADA}.", null);
        }

        var nomeArquivo = $"SCC_DAT_CORR_{dataProcessamento}_Carga";
        var caminhoArquivo = $"{diretorioTemporario}/{nomeArquivo}";

        var stringBuilder = new StringBuilder();

        MontarCabecalho(stringBuilder, dataProcessamento);
        var linhasNaoProcessadas = await ProcessarCorpo(stringBuilder, arquivoEntrada, dataProcessamento);
        sequencialRegistro++;
        MontarRodape(stringBuilder);

        using (var arquivo = new StreamWriter(caminhoArquivo, false, Encoding.UTF8))
        {
            await arquivo.WriteAsync(stringBuilder.ToString());
        };

        var bytesArquivo = File.ReadAllBytes(caminhoArquivo);
        const string CONTENT_TYPE = "application/...";
        var arquivoSaida = new FileContentResult(bytesArquivo, CONTENT_TYPE) { FileDownloadName = nomeArquivo };

        return (arquivoSaida, null, linhasNaoProcessadas);
    }

    private async Task<List<int>?> ProcessarCorpo(StringBuilder stringBuilder, IFormFile arquivoEntrada, string dataProcessamento)
    {
        const char CARACTERE_SEPARADOR = ';';

        const int COLUNA_CPF = 0;
        const int COLUNA_NOME = 3;
        const int COLUNA_CONTRATO = 4;
        const int COLUNA_AGENCIA = 5;
        const int COLUNA_CONTA = 6;
        const int COLUNA_VALOR = 9;
        const int COLUNA_SEGURADORA = 11;
        const int COLUNA_TIPO_PAGAMENTO = 12;
        const int COLUNA_DATA_LIBERACAO = 13;
        const int COLUNA_DATA_VENCIMENTO = 14;
        const int COLUNA_REGRA_SEGURO = 15;
        const int COLUNA_TAXA = 16;

        var linhasNaoProcessadas = new List<int>();

        using var stream = new MemoryStream();
        await arquivoEntrada.CopyToAsync(stream);

        using var textFieldParser = new TextFieldParser(stream);
        textFieldParser.TextFieldType = FieldType.Delimited;
        textFieldParser.SetDelimiters(CARACTERE_SEPARADOR.ToString());

        var numeroLinha = 0;
        stream.Seek(0, SeekOrigin.Begin);
        while (!textFieldParser.EndOfData)
        {
            numeroLinha++;
            sequencialRegistro++;

            var linha = textFieldParser.ReadFields();
            if (linha == null) { continue; }

            try 
            {
                var seguro = new Seguro
                (
                    linha.GetValue(COLUNA_CPF)?.ToString()!,
                    linha.GetValue(COLUNA_NOME)?.ToString()!,
                    linha.GetValue(COLUNA_CONTRATO)?.ToString()!,
                    linha.GetValue(COLUNA_AGENCIA)?.ToString()!,
                    linha.GetValue(COLUNA_CONTA)?.ToString()!,
                    linha.GetValue(COLUNA_VALOR)?.ToString()!,
                    linha.GetValue(COLUNA_SEGURADORA)?.ToString()!,
                    linha.GetValue(COLUNA_TIPO_PAGAMENTO)?.ToString()!,
                    linha.GetValue(COLUNA_DATA_LIBERACAO)?.ToString()!,
                    linha.GetValue(COLUNA_DATA_VENCIMENTO)?.ToString()!,
                    linha.GetValue(COLUNA_REGRA_SEGURO)?.ToString()!,
                    linha.GetValue(COLUNA_TAXA)?.ToString()!
                );
                MontarCorpo(stringBuilder, seguro, dataProcessamento);
            }
            catch
            {
                linhasNaoProcessadas.Add(numeroLinha);
            }
        }
        return linhasNaoProcessadas;
    }

    private void MontarCabecalho(StringBuilder stringBuilder, string dataProcessamento)
    {
        const char TIPO_REGISTRO = '0';
        const string NOME_ARQUIVO = "CORRETORA";

        stringBuilder.AppendLine
        (
            TIPO_REGISTRO +
            sequencialRegistro.ToString().PadLeft(6, CARACTERE_COMPLEMENTAR) +
            NOME_ARQUIVO +
            dataProcessamento +
            CODIGO_BANCO +
            string.Empty.PadLeft(147, CARACTERE_COMPLEMENTAR_ALFANUMERICO)
        );
    }

    private void MontarCorpo(StringBuilder stringBuilder, Seguro seguro, string dataProcessamento)
    {
        const char TIPO_REGISTRO = '1';
        const char STATUS_CONTRATO = '0';
        const char TIPO_SEGURO = '3';
        const char ORIGEM_CONTRATACAO = 'A';
        const char PARCELA_UNICA = 'S';
        const char FLAG_ADITAMENTO = 'N';

        var quantidadeParcelas = ((seguro.DataVencimento.Year - seguro.DataLiberacao.Year) * 12) + seguro.DataVencimento.Month - seguro.DataLiberacao.Month;
        var valorPremio = seguro.Taxa * seguro.Valor * quantidadeParcelas;

        stringBuilder.AppendLine
        (
            TIPO_REGISTRO +
            sequencialRegistro.ToString().PadLeft(6, CARACTERE_COMPLEMENTAR) +
            seguro.Contrato.PadLeft(13, CARACTERE_COMPLEMENTAR) +
            string.Empty.PadLeft(6, CARACTERE_COMPLEMENTAR) +
            string.Empty.PadLeft(6, CARACTERE_COMPLEMENTAR) +
            CODIGO_BANCO.PadLeft(3, CARACTERE_COMPLEMENTAR) +
            seguro.Agencia.PadLeft(6, CARACTERE_COMPLEMENTAR) +
            seguro.Conta.PadLeft(20, CARACTERE_COMPLEMENTAR) +
            $"{seguro.Cpf[..9]}0000{seguro.Cpf[9..]}00" +
            seguro.Nome.PadRight(35, CARACTERE_COMPLEMENTAR_ALFANUMERICO) +
            dataProcessamento +
            ((int)(seguro.Valor * 100)).ToString().PadLeft(10, CARACTERE_COMPLEMENTAR) +
            seguro.DataLiberacao.ToString(FORMATO_DATA_CONVERSAO) +
            seguro.DataVencimento.ToString(FORMATO_DATA_CONVERSAO) +
            string.Empty.PadLeft(8, CARACTERE_COMPLEMENTAR) +
            STATUS_CONTRATO +
            TIPO_SEGURO +
            string.Empty.PadLeft(2, CARACTERE_COMPLEMENTAR) +
            quantidadeParcelas.ToString().PadLeft(3, CARACTERE_COMPLEMENTAR) +
            ORIGEM_CONTRATACAO +
            seguro.RegraSeguro.PadLeft(6, CARACTERE_COMPLEMENTAR) +
            ((int)(seguro.Taxa * 10000)).ToString().PadLeft(10, CARACTERE_COMPLEMENTAR) +
            ((int)valorPremio).ToString().PadLeft(10, CARACTERE_COMPLEMENTAR) +
            PARCELA_UNICA +
            string.Empty.PadLeft(3, CARACTERE_COMPLEMENTAR) +
            FLAG_ADITAMENTO +
            string.Empty.PadLeft(3, CARACTERE_COMPLEMENTAR) +
            string.Empty.PadLeft(8, CARACTERE_COMPLEMENTAR) +
            string.Empty.PadLeft(3, CARACTERE_COMPLEMENTAR) +
            string.Empty.PadLeft(10, CARACTERE_COMPLEMENTAR) +
            string.Empty.PadLeft(10, CARACTERE_COMPLEMENTAR) +
            string.Empty.PadLeft(8, CARACTERE_COMPLEMENTAR)
        );
    }

    private void MontarRodape(StringBuilder stringBuilder)
    {
        const char TIPO_REGISTRO = '9';

        stringBuilder.Append
        (
            TIPO_REGISTRO +
            sequencialRegistro.ToString().PadLeft(6, CARACTERE_COMPLEMENTAR) +
            string.Empty.PadLeft(167, CARACTERE_COMPLEMENTAR_ALFANUMERICO)
        );
    }
}
