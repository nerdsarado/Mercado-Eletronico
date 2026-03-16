using System;

namespace MercadoEletronico.Models
{
    public class Cotacao
    {
        public string NumeroCotacao { get; set; } = "";
        public string Portal { get; set; } = "Mercado eletrônico";
        public string Empresa { get; set; } = "Nestle";
        public DateTime DataVencimento { get; set; }
        public TimeSpan HorarioVencimento { get; set; }
        public DateTime DataRegistro { get; set; }
        public TimeSpan HorarioRegistro { get; set; }
        public string Solicitante { get; set; } = "";
        public string Vendedor { get; set; } = "";
        public string Status { get; set; } = "";
        public string IDTransacao { get; set; } = "";
        public List<string> Itens { get; set; }
    }
}