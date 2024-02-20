using System.Globalization;

namespace migracao_seguros.Entidades;

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
                  DateTime dataLiberacao,
                  DateTime dataVencimento,
                  string regraSeguro,
                  string taxa)
    {
        Cpf = cpf;
        Nome = nome;
        Contrato = contrato;
        Agencia = agencia;
        Conta = conta;
        Valor = double.Parse(valor, CultureInfo.InvariantCulture);
        Seguradora = seguradora;
        TipoPagamento = tipoPagamento;
        DataLiberacao = dataLiberacao;
        DataVencimento = dataVencimento;
        RegraSeguro = regraSeguro;
        Taxa = double.Parse(taxa, CultureInfo.InvariantCulture);
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