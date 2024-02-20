using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Mvc;
using migracao_seguros.Entidades;

namespace migracao_seguros.Services;

public class ArquivoService : IArquivoService
{
    private readonly string diretorioTemporario = Path.GetTempPath();
    private const char CARACTERE_COMPLEMENTAR = '0';
    private const char CARACTERE_COMPLEMENTAR_ALFANUMERICO = ' ';
    private const string CODIGO_BANCO = "047";
    private const string FORMATO_DATA_CONVERSAO = "yyyyMMdd";

    [HttpPost("processar")]
    public (FileContentResult? arquivoConvertido, string? erro) Processar(IFormFile arquivoRecebido)
    {
        var dataProcessamento = DateTime.Now;

        if(arquivoRecebido == null || arquivoRecebido.Length == 0)
        {
            return (null, "Um arquivo xlsx deve ser selecionado.");
        }

        const string EXTENSAO_XLSX = ".xlsx";
        var extensao = arquivoRecebido.FileName.Substring(arquivoRecebido.FileName.Length - EXTENSAO_XLSX.Length);
        if(extensao != EXTENSAO_XLSX)
        {
            return (null, "O arquivo enviado deve ser do tipo xlsx.");
        }

        try 
        {
            var seguros = LerArquivo(arquivoRecebido);
            if(seguros == null)
            {
                return (null, "Ocorreu um erro ao ler o arquivo.");
            }

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

            var bytesArquivo = File.ReadAllBytes(caminhoArquivo);
            const string CONTENT_TYPE = "application/...."; 
            return (new FileContentResult(bytesArquivo, CONTENT_TYPE) {FileDownloadName = nomeArquivo}, null);
        }
        catch(Exception)
        {
            return (null, "Verifique o formato do arquivo e tente novamente.");
        }
        
    }

    private static IList<Seguro>? LerArquivo(IFormFile arquivo)
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
        arquivo.CopyTo(stream);
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

    private static string ObterValorDaCelula(WorkbookPart workbookPart, Row linha, int coluna)
    {
        var cell = linha.Elements<Cell>().ElementAt(coluna);
        
        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
        {
            int sharedStringIndex = int.Parse(cell.InnerText);
            return workbookPart.SharedStringTablePart!.SharedStringTable.Elements<SharedStringItem>().ElementAt(sharedStringIndex).InnerText;
        }
        
        return cell.InnerText;
    }

    private static DateTime ObterDataDaCelula(WorkbookPart workbookPart, Row linha, int coluna)
    {
        var valor = ObterValorDaCelula(workbookPart, linha, coluna);
        return DateTime.FromOADate(double.Parse(valor));
    }

    private static void MontarCabecalho(StreamWriter arquivo, int sequencialRegistro, DateTime dataProcessamento)
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

    private static void MontarCorpo(StreamWriter arquivo, int sequencial, Seguro seguro)
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

    private static void MontarRodape(StreamWriter arquivo, int sequencialRegistro)
    {
        const char TIPO_REGISTRO = '9';

        arquivo.Write(TIPO_REGISTRO);
        arquivo.Write(sequencialRegistro.ToString().PadLeft(6, CARACTERE_COMPLEMENTAR));
        arquivo.WriteLine(string.Empty.PadLeft(167, CARACTERE_COMPLEMENTAR_ALFANUMERICO));
    }
}