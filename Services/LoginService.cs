using Microsoft.Playwright;
using MercadoEletronico.Models;
using System.Threading.Tasks;
using System;

namespace MercadoEletronico.Services
{
    public class LoginService
    {
        private readonly IPage _page;

        public LoginService(IPage page)
        {
            _page = page;
        }

        public async Task RealizarLoginAsync(int i)
        {
            var contas = new Models.Contas();
            Console.WriteLine("\n1. FAZENDO LOGIN...");
            Console.WriteLine("   🌐 Navegando para página de login...");

            try
            {
                    // Opção 1: Usar timeout maior
                    await _page.GotoAsync(Configuracao.URL_LOGIN, new PageGotoOptions
                    {
                        WaitUntil = WaitUntilState.DOMContentLoaded, // Alterado de NetworkIdle
                        Timeout = 60000
                    });

                    // Opção 2: Esperar por elemento específico da página de login
                    await _page.WaitForSelectorAsync("#LoginName", new PageWaitForSelectorOptions
                    {
                        Timeout = 30000
                    });

                    //await CapturarScreenshotAsync("00_antes_login.png");

                    Console.WriteLine("   ⌨️  Preenchendo credenciais...");
                    await _page.FillAsync("#LoginName", contas.Usuarios[i]);
                    await _page.FillAsync("#RAWSenha", contas.Senhas[i]);

                    Console.WriteLine("   🔘 Clicando em 'Entrar'...");
                    await _page.ClickAsync("#SubmitAuth");

                    // Aguardar por elemento indicativo de login bem-sucedido
                    await Task.Delay(Configuracao.TIMEOUT_MEDIO * 2);

                    // Tentar diferentes estratégias para confirmar login
                    bool loginSucesso = await VerificarLoginSucessoAsync();

                    if (!loginSucesso)
                    {
                        Console.WriteLine("   ⚠️  Login pode não ter sido bem-sucedido, continuando...");
                    }

                    Console.WriteLine($"   ✅ Tentativa de login concluída! URL: {_page.Url}");
                    //await CapturarScreenshotAsync("01_apos_login.png");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"   ❌ Erro durante login: {ex.Message}");
                //await CapturarScreenshotAsync("erro_login.png");
                throw; // Re-lançar para que o fluxo principal saiba do erro
            }
        }

        private async Task<bool> VerificarLoginSucessoAsync()
        {
            try
            {
                // Aguardar um pouco mais para carregamento
                await Task.Delay(5000);

                // Verificar se estamos em uma URL diferente da de login
                if (!_page.Url.Contains("Login", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }

                // Verificar se há elementos indicativos de dashboard/logado
                var elementosIndicadores = new[]
                {
                    ".dashboard", ".me-sidebar", ".sidebar", ".user-menu",
                    "[data-cy='user-menu']", ".user-profile"
                };

                foreach (var selector in elementosIndicadores)
                {
                    try
                    {
                        var element = await _page.QuerySelectorAsync(selector);
                        if (element != null && await element.IsVisibleAsync())
                        {
                            return true;
                        }
                    }
                    catch { continue; }
                }

                // Verificar por texto indicativo
                var pageContent = await _page.ContentAsync();
                if (pageContent.Contains("Bem-vindo", StringComparison.OrdinalIgnoreCase) ||
                    pageContent.Contains("Dashboard", StringComparison.OrdinalIgnoreCase) ||
                    pageContent.Contains("Menu", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }

                return false;
            }
            catch
            {
                return false;
            }
        }

        private async Task CapturarScreenshotAsync(string nomeArquivo)
        {
            try
            {
                await _page.ScreenshotAsync(new PageScreenshotOptions
                {
                    Path = nomeArquivo,
                    FullPage = true
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"   ⚠️  Erro ao capturar screenshot: {ex.Message}");
            }
        }
    }
}