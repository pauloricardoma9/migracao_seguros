using System.Globalization;

namespace migracao_seguros.Models;

public class Seguro
{
    public Seguro(string cpf,
                  string nome,
                  string contrato,
                  string agencia,
                  string conta,
                  string valor,
                  string seguradora,
                  string tipoPagamento,
                  string dataLiberacao,
                  string dataVencimento,
                  string regraSeguro,
                  string taxa)
    {
        const string FORMATO_DATA = "dd/MM/yyyy";
        
        Cpf = cpf;
        Nome = nome[..Math.Min(nome.Length, 35)];
        Contrato = contrato;
        Agencia = agencia;
        Conta = conta;
        Valor = double.Parse(valor.Replace("R$", ""));
        Seguradora = seguradora;
        TipoPagamento = tipoPagamento;
        DataLiberacao = DateTime.ParseExact(dataLiberacao, FORMATO_DATA, CultureInfo.InvariantCulture);;
        DataVencimento = DateTime.ParseExact(dataVencimento, FORMATO_DATA, CultureInfo.InvariantCulture);;
        RegraSeguro = regraSeguro;
        Taxa = double.Parse(taxa);
    }

    public string Cpf { get; private set; }
    public string Nome { get; private set; }
    public string Contrato { get; private set; }
    public string Agencia { get; private set; }
    public string Conta { get; private set; }
    public double Valor { get; private set; }
    public string Seguradora { get; private set; }
    public string TipoPagamento { get; private set; }
    public DateTime DataLiberacao { get; private set; }
    public DateTime DataVencimento { get; private set; }
    public string RegraSeguro { get; private set; }
    public double Taxa { get; private set; }
}