using System.Data;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Mvc;
using migracao_seguros.Entidades;

namespace migracao_seguros.Controllers;

[ApiController]
[Route("[controller]")]
public class ArquivoController : ControllerBase
{
    private readonly ILogger<ArquivoController> _logger;
    private readonly string diretorioTemporario = Path.GetTempPath();
    private const char CARACTERE_COMPLEMENTAR = '0';
    private const char CARACTERE_COMPLEMENTAR_ALFANUMERICO = ' ';
    private const string CODIGO_BANCO = "047";
    private const string FORMATO_DATA_CONVERSAO = "yyyyMMdd";

    public ArquivoController(ILogger<ArquivoController> logger)
    {
        _logger = logger;
    }

    [HttpPost("Processar")]
    public IActionResult Processar(IFormFile arquivoXlsx)
    {
        var dataProcessamento = DateTime.Now;

        // Verificar se foi feito o upload do arquivo original
        if (arquivoXlsx == null || arquivoXlsx.Length == 0)
        {
            return BadRequest("O arquivo não foi enviado.");
        }

        // Ler arquivo original
        var seguros = LerArquivoOriginal(arquivoXlsx);
        if(seguros == null)
        {
            return BadRequest("Ocorreu um erro ao ler o arquivo.");
        }

        // Criar arquivo
        var nomeArquivo = "NomeArquivo";
        var caminhoArquivo = $"{diretorioTemporario}/{nomeArquivo}";

        var arquivo = new StreamWriter(caminhoArquivo, false, Encoding.UTF8);

        var sequencialRegistro = 1;
        MontarCabecalho(arquivo, sequencialRegistro, dataProcessamento);

        foreach(var seguro in seguros)
        {
            sequencialRegistro++;
            MontarCorpo(arquivo, sequencialRegistro, seguro);
        }

        sequencialRegistro++;
        MontarRodape(arquivo, sequencialRegistro);

        arquivo.Close();

        // Retornar o arquivo
        var bytesArquivo = System.IO.File.ReadAllBytes(caminhoArquivo);
        const string CONTENT_TYPE = "application/...."; 
        return new FileContentResult(bytesArquivo, CONTENT_TYPE) {FileDownloadName = nomeArquivo};
    }

    private IList<Seguro>? LerArquivoOriginal(IFormFile arquivoOriginal)
    {
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

        using var stream = new MemoryStream();
        arquivoOriginal.CopyTo(stream);
        using var spreadsheetDocument = SpreadsheetDocument.Open(stream, false);
        var workbookPart = spreadsheetDocument.WorkbookPart;
        if (workbookPart == null)
        {
            return null;
        }

        var worksheetPart = workbookPart.WorksheetParts.First();
        var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

        var seguros = new List<Seguro>();
        foreach (var linha in sheetData.Elements<Row>().Skip(1))
        {
            var seguro = new Seguro
            (
                ObterValorDaCelula(workbookPart, linha, COLUNA_CPF),
                ObterValorDaCelula(workbookPart, linha, COLUNA_NOME),
                ObterValorDaCelula(workbookPart, linha, COLUNA_CONTRATO),
                ObterValorDaCelula(workbookPart, linha, COLUNA_AGENCIA),
                ObterValorDaCelula(workbookPart, linha, COLUNA_CONTA),
                ObterValorDaCelula(workbookPart, linha, COLUNA_VALOR),
                ObterValorDaCelula(workbookPart, linha, COLUNA_SEGURADORA),
                ObterValorDaCelula(workbookPart, linha, COLUNA_TIPO_PAGAMENTO),
                ObterDataDaCelula(workbookPart, linha, COLUNA_DATA_LIBERACAO),
                ObterDataDaCelula(workbookPart, linha, COLUNA_DATA_VENCIMENTO),
                ObterValorDaCelula(workbookPart, linha, COLUNA_REGRA_SEGURO),
                ObterValorDaCelula(workbookPart, linha, COLUNA_TAXA)
            );
            seguros.Add(seguro);
        }

        return seguros;
    }

    private string ObterValorDaCelula(WorkbookPart workbookPart, Row linha, int coluna)
    {
        var cell = linha.Elements<Cell>().ElementAt(coluna);
        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
        {
            // Se o tipo de dado for SharedString, obter o valor a partir do índice na tabela SharedStringTable
            int sharedStringIndex = int.Parse(cell.InnerText);
            return workbookPart.SharedStringTablePart!.SharedStringTable.Elements<SharedStringItem>().ElementAt(sharedStringIndex).InnerText;
        }
        else
        {
            // Se não for SharedString, obter o valor diretamente
            return cell.InnerText;
        }
    }

    private DateTime ObterDataDaCelula(WorkbookPart workbookPart, Row linha, int coluna)
    {
        var valor = ObterValorDaCelula(workbookPart, linha, coluna);
        return DateTime.FromOADate(double.Parse(valor));
    }

    private void MontarCabecalho(StreamWriter arquivo, int sequencialRegistro, DateTime dataProcessamento)
    {
        const char TIPO_REGISTRO = '0';
        const string NOME_ARQUIVO = "CORRETORA";

        arquivo.Write(TIPO_REGISTRO);
        arquivo.Write(sequencialRegistro.ToString().PadLeft(6, CARACTERE_COMPLEMENTAR));
        arquivo.Write(NOME_ARQUIVO);
        arquivo.Write(dataProcessamento.ToString(FORMATO_DATA_CONVERSAO));
        arquivo.Write(CODIGO_BANCO);
        arquivo.WriteLine(string.Empty.PadLeft(147, CARACTERE_COMPLEMENTAR_ALFANUMERICO));
    }

    private void MontarCorpo(StreamWriter arquivo, int sequencial, Seguro seguro)
    {
        const char TIPO_REGISTRO = '1';
        const char STATUS_CONTRATO = '0';
        const char TIPO_SEGURO = '3';
        const char ORIGEM_CONTRATACAO = 'A';
        const char PARCELA_UNICA = 'S';
        const char FLAG_ADITAMENTO = 'N';

        var quantidadeParcelas = ((seguro.DataVencimento.Year - seguro.DataLiberacao.Year) * 12) + seguro.DataVencimento.Month - seguro.DataLiberacao.Month;
        var valorPremio = seguro.Taxa * seguro.Valor * quantidadeParcelas;

        arquivo.Write(TIPO_REGISTRO);
        arquivo.Write(sequencial.ToString().PadLeft(6, CARACTERE_COMPLEMENTAR));
        arquivo.Write(seguro.Contrato.PadLeft(13, CARACTERE_COMPLEMENTAR));
        arquivo.Write(string.Empty.PadLeft(6, CARACTERE_COMPLEMENTAR));
        arquivo.Write(string.Empty.PadLeft(6, CARACTERE_COMPLEMENTAR));
        arquivo.Write(CODIGO_BANCO.PadLeft(3, CARACTERE_COMPLEMENTAR));
        arquivo.Write(seguro.Agencia.PadLeft(6, CARACTERE_COMPLEMENTAR));
        arquivo.Write(seguro.Conta.PadLeft(20, CARACTERE_COMPLEMENTAR));
        arquivo.Write($"{seguro.Cpf[..9]}0000{seguro.Cpf[9..]}00");
        arquivo.Write(seguro.Nome.PadRight(35, CARACTERE_COMPLEMENTAR_ALFANUMERICO));
        arquivo.Write(string.Empty.PadLeft(8, CARACTERE_COMPLEMENTAR));
        arquivo.Write(((int)(seguro.Valor * 100)).ToString().PadLeft(10, CARACTERE_COMPLEMENTAR));
        arquivo.Write(seguro.DataLiberacao.ToString(FORMATO_DATA_CONVERSAO));
        arquivo.Write(seguro.DataVencimento.ToString(FORMATO_DATA_CONVERSAO));
        arquivo.Write(string.Empty.PadLeft(8, CARACTERE_COMPLEMENTAR));
        arquivo.Write(STATUS_CONTRATO);
        arquivo.Write(TIPO_SEGURO);
        arquivo.Write(string.Empty.PadLeft(2, CARACTERE_COMPLEMENTAR));
        arquivo.Write(quantidadeParcelas.ToString().PadLeft(3, CARACTERE_COMPLEMENTAR));
        arquivo.Write(ORIGEM_CONTRATACAO);
        arquivo.Write(seguro.RegraSeguro.PadLeft(6, CARACTERE_COMPLEMENTAR));
        arquivo.Write(((int)(seguro.Taxa * 10000)).ToString().PadLeft(10, CARACTERE_COMPLEMENTAR));
        arquivo.Write(((int)valorPremio).ToString().PadLeft(10, CARACTERE_COMPLEMENTAR));
        arquivo.Write(PARCELA_UNICA);
        arquivo.Write(string.Empty.PadLeft(3, CARACTERE_COMPLEMENTAR));
        arquivo.Write(FLAG_ADITAMENTO);
        arquivo.Write(string.Empty.PadLeft(3, CARACTERE_COMPLEMENTAR));
        arquivo.Write(string.Empty.PadLeft(8, CARACTERE_COMPLEMENTAR));
        arquivo.Write(string.Empty.PadLeft(3, CARACTERE_COMPLEMENTAR));
        arquivo.Write(string.Empty.PadLeft(10, CARACTERE_COMPLEMENTAR));
        arquivo.Write(string.Empty.PadLeft(10, CARACTERE_COMPLEMENTAR));
        arquivo.WriteLine(string.Empty.PadLeft(8, CARACTERE_COMPLEMENTAR));
    }

    private void MontarRodape(StreamWriter arquivo, int sequencialRegistro)
    {
        const char TIPO_REGISTRO = '9';

        arquivo.Write(TIPO_REGISTRO);
        arquivo.Write(sequencialRegistro.ToString().PadLeft(6, CARACTERE_COMPLEMENTAR));
        arquivo.WriteLine(string.Empty.PadLeft(167, CARACTERE_COMPLEMENTAR_ALFANUMERICO));
    }
}
