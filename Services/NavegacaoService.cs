using MercadoEletronico.Models;
using Microsoft.Playwright;
using System;
using System.Threading.Tasks;

namespace MercadoEletronico.Services
{
    public class NavegacaoService
    {
        private readonly IPage _page;

        public NavegacaoService(IPage page)
        {
            _page = page;
        }

        public async Task NavegarParaOportunidadesAsync()
        {
            Console.WriteLine("\n4. NAVEGANDO PARA 'OPORTUNIDADES A RESPONDER'...");

            // Primeiro verificar se já estamos na página correta E se as cotações estão visíveis
            await VerificarSeCotacoesEstaoVisiveisAsync();

            // Tentar clicar no item específico do menu
            Console.WriteLine("   🖱️  Tentando clicar em 'Oportunidades a Responder' no menu lateral...");

            // Usar o seletor exato que você forneceu
            var sucesso = await TentarClicarComSeletorExatoAsync();

            if (!sucesso)
            {
                sucesso = await TentarOutrosSeletoresAsync();
            }

            if (!sucesso)
            {
                sucesso = await TentarViaTextoExatoAsync();
            }

            if (!sucesso)
            {
                sucesso = await TentarViaJavaScriptAsync();
            }

            if (sucesso)
            {
                await AguardarCarregamentoCotacoesAsync();
                return;
            }

            Console.WriteLine("   ⚠️  Não foi possível clicar no menu, mas continuando...");
        }

        private async Task<bool> TentarClicarComSeletorExatoAsync()
        {
            try
            {
                Console.WriteLine("   📌 Tentando seletor exato: span.me-sidebar-cell-simple-item__label...");

                // Aguardar o elemento estar disponível
                await _page.WaitForSelectorAsync("span.me-sidebar-cell-simple-item__label", new PageWaitForSelectorOptions
                {
                    Timeout = 5000
                });

                // Buscar todos os elementos com essa classe
                var elementos = await _page.QuerySelectorAllAsync("span.me-sidebar-cell-simple-item__label");
                Console.WriteLine($"   📊 Encontrados {elementos.Count} elementos com a classe");

                foreach (var elemento in elementos)
                {
                    try
                    {
                        var texto = await elemento.InnerTextAsync();
                        Console.WriteLine($"   🔍 Texto do elemento: '{texto}'");

                        if (texto.Contains("Oportunidades a Responder", StringComparison.OrdinalIgnoreCase))
                        {
                            Console.WriteLine($"   ✅ Elemento correto encontrado!");

                            // Verificar se está visível
                            if (await elemento.IsVisibleAsync())
                            {
                                // Scroll para o elemento
                                await elemento.ScrollIntoViewIfNeededAsync();
                                await Task.Delay(1000);

                                // Clicar
                                await elemento.ClickAsync();
                                Console.WriteLine("   🖱️  Clicado no menu 'Oportunidades a Responder'!");
                                return true;
                            }
                            else
                            {
                                Console.WriteLine("   ⚠️  Elemento encontrado mas não está visível");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"   ⚠️  Erro ao verificar elemento: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"   ⚠️  Erro ao buscar seletor exato: {ex.Message}");
            }

            return false;
        }

        private async Task<bool> TentarOutrosSeletoresAsync()
        {
            var seletores = new[]
            {
                "[data-cy='me-sidebar-cell-simple-item__label']",
                ".me-sidebar-cell-simple-item__label",
                ".me-sidebar-cell-simple-item",
                ".me-sidebar-item",
                "[data-cy*='sidebar']",
                ".sidebar-item",
                "nav li",
                "aside li"
            };

            foreach (var selector in seletores)
            {
                try
                {
                    Console.WriteLine($"   🔍 Tentando seletor: {selector}");
                    var elementos = await _page.QuerySelectorAllAsync(selector);

                    foreach (var elemento in elementos)
                    {
                        var texto = await elemento.InnerTextAsync();
                        if (texto.Contains("Oportunidades a Responder", StringComparison.OrdinalIgnoreCase))
                        {
                            await elemento.ScrollIntoViewIfNeededAsync();
                            await Task.Delay(1000);
                            await elemento.ClickAsync();
                            Console.WriteLine($"   ✅ Clicado usando seletor: {selector}");
                            return true;
                        }
                    }
                }
                catch { continue; }
            }

            return false;
        }

        private async Task<bool> TentarViaTextoExatoAsync()
        {
            try
            {
                Console.WriteLine("   🔍 Tentando clicar pelo texto exato...");

                // Usar a função de clique por texto do Playwright
                await _page.ClickAsync("text='Oportunidades a Responder'", new PageClickOptions
                {
                    Timeout = 5000
                });

                Console.WriteLine("   ✅ Clicado via texto!");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"   ⚠️  Não foi possível clicar pelo texto: {ex.Message}");
                return false;
            }
        }

        private async Task<bool> TentarViaJavaScriptAsync()
        {
            try
            {
                Console.WriteLine("   💻 Usando JavaScript para encontrar e clicar...");

                var encontrado = await _page.EvaluateAsync<bool>(@"
                    () => {
                        // Buscar por texto exato
                        const elementos = Array.from(document.querySelectorAll('*'));
                        const alvo = elementos.find(el => {
                            const texto = el.textContent || '';
                            return texto.trim() === 'Oportunidades a Responder';
                        });
                        
                        if (alvo) {
                            console.log('Elemento encontrado via JS:', alvo);
                            alvo.click();
                            return true;
                        }
                        
                        // Tentar por classe específica
                        const elementosClasse = document.querySelectorAll('.me-sidebar-cell-simple-item__label');
                        for (const el of elementosClasse) {
                            if (el.textContent.includes('Oportunidades a Responder')) {
                                el.click();
                                return true;
                            }
                        }
                        
                        return false;
                    }
                ");

                if (encontrado)
                {
                    Console.WriteLine("   ✅ Clicado via JavaScript!");
                    return true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"   ⚠️  Erro no JavaScript: {ex.Message}");
            }

            return false;
        }

        private async Task VerificarSeCotacoesEstaoVisiveisAsync()
        {
            try
            {
                Console.WriteLine("   👀 Verificando se cotações já estão visíveis...");

                // Verificar se já existem elementos de cotações na página
                var cotacoes = await _page.QuerySelectorAllAsync("div.left-info, div.doc-title, .cotacao, [class*='quote'], [class*='cotacao']");

                if (cotacoes.Count > 0)
                {
                    Console.WriteLine($"   📊 Já existem {cotacoes.Count} elementos de cotação na página");

                    // Verificar se algum contém "Cotação" no texto
                    foreach (var cotacao in cotacoes)
                    {
                        try
                        {
                            var texto = await cotacao.InnerTextAsync();
                            if (texto.Contains("Cotação", StringComparison.OrdinalIgnoreCase))
                            {
                                Console.WriteLine("   ✅ Cotações já estão carregadas na página!");
                                return;
                            }
                        }
                        catch { }
                    }
                }

                Console.WriteLine("   ℹ️  Cotações não estão visíveis, precisa clicar no menu");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"   ⚠️  Erro ao verificar cotações: {ex.Message}");
            }
        }

        private async Task AguardarCarregamentoCotacoesAsync()
        {
            Console.WriteLine("   ⏳ Aguardando carregamento das cotações...");

            try
            {
                // Aguardar por elementos de cotação
                await _page.WaitForSelectorAsync("div.left-info, div.doc-title, .cotacao-item, [class*='quote']", new PageWaitForSelectorOptions
                {
                    Timeout = 10000
                });

                await Task.Delay(Configuracao.TIMEOUT_MEDIO);
                Console.WriteLine($"   📍 URL atual após clique: {_page.Url}");

                // Capturar screenshot
                await _page.ScreenshotAsync(new PageScreenshotOptions
                {
                    Path = "04_menu_clicado.png",
                    FullPage = true
                });

                Console.WriteLine("   ✅ Menu clicado e página carregada!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"   ⚠️  Aguardando carregamento: {ex.Message}");
                // Mesmo se não encontrar, dar um tempo
                await Task.Delay(3000);
            }
        }
    }
}