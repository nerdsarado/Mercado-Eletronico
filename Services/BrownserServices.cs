using Microsoft.Playwright;
using System.Threading.Tasks;

namespace MercadoEletronico.Services
{
    public class BrowserService
    {
        public IBrowser _browser;
        public async Task<IPage> InicializarBrowserAsync()
        {
            Console.WriteLine("=== Sistema de Automação Mercado Eletrônico ===");
            Console.WriteLine("\n🖥️  Configurando navegador em tela cheia...");

            var playwright = await Playwright.CreateAsync();

            _browser = await playwright.Chromium.LaunchAsync(new BrowserTypeLaunchOptions
            {
                ExecutablePath = @"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
                Headless = true,
                SlowMo = 300
            });

            var context = await _browser.NewContextAsync();
            var page = await context.NewPageAsync();

            Console.WriteLine("✅ Navegador configurado!");
            return page;
        }
    }
}