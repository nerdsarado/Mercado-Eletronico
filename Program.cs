using MercadoEletronico.Services;
using Microsoft.Playwright;
using System;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Linq;

namespace MercadoEletronico
{
    class Program
    {
        static async Task Main(string[] args)
        {
            Console.WriteLine("🚀 Iniciando automação Mercado Eletrônico");
            Console.WriteLine("=========================================\n");

            PlanilhaControleService planilhaControle = null;

            try
            {
                // 1. Inicializar serviço de planilha
                planilhaControle = new PlanilhaControleService();
                var cotacoesProcessadas = planilhaControle.CarregarCotacoesProcessadas();

                Console.WriteLine($"📊 Cotações já processadas: {cotacoesProcessadas.Count}\n");

                // 2. Inicializar navegador
                var browserService = new BrowserService();
                var page = await browserService.InicializarBrowserAsync();

                // 3. Serviços
                var loginService = new LoginService(page);
                var modalHandler = new ModalHandlerService(page);
                var navegacaoService = new NavegacaoService(page);
                var coletaService = new ColetaService(page, modalHandler, cotacoesProcessadas, planilhaControle);
                var exportadorService = new ExportadorService();
                var contas = new Models.Contas();

                await coletaService.Executar(page, loginService,
                                             modalHandler, navegacaoService, 
                                             coletaService, cotacoesProcessadas, exportadorService);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}