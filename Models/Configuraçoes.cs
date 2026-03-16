namespace MercadoEletronico.Models
{
    public static class Configuracao
    {
        public const string URL_LOGIN = "https://me.com.br/do/Login.mvc/LoginNew";
        public const int TIMEOUT_PADRAO = 30000;
        public const int TIMEOUT_CURTO = 2000;
        public const int TIMEOUT_MEDIO = 3000;
    }
    public class Contas
    {
        public List<string> Empresas { get; set; } = new List<string> { "Uniao", "Ventura", "Alianca" };
        public List<string> Usuarios { get; set; } = new List<string> { "INFONESTLE" , "A76FEB51"};
        public List<string> Senhas { get; set; } = new List<string> { "Ventura2025*" , "Nestle!@24"};

    }
}