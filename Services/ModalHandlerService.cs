using MercadoEletronico.Models;
using Microsoft.Playwright;
using System;
using System.Threading.Tasks;

namespace MercadoEletronico.Services
{
    public class ModalHandlerService
    {
        private readonly IPage _page;

        public ModalHandlerService(IPage page)
        {
            _page = page;
        }

        public async Task LidarComTodosModaisAsync()
        {
            await Task.Delay(Configuracao.TIMEOUT_MEDIO);

            Console.WriteLine("\n2. VERIFICANDO MODAIS...");

            // Primeiro modal
            await VerificarEClicarVerDepoisAsync();

            // Segundo modal
            await VerificarEClicarOkEntendiAsync();

            // Terceiro modal (checkboxes e botões de aceite)
            await VerificarEClicarCheckboxCienteAsync();
            await VerificarEClicarBotaoGravarAsync();

            // Verificação final
            await VerificarModaisRestantesAsync();
        }

        private async Task VerificarEClicarVerDepoisAsync()
        {
            Console.WriteLine("   🔍 Verificando se existe botão 'Ver depois'...");

            try
            {
                await Task.Delay(Configuracao.TIMEOUT_CURTO);
                var verDepoisButton = await _page.QuerySelectorAsync("button.see-later-btn");

                if (verDepoisButton != null && await verDepoisButton.IsVisibleAsync())
                {
                    Console.WriteLine("   ✅ Botão 'Ver depois' encontrado e visível!");
                    //await CapturarScreenshotModalAsync("modal1_ver_depois_antes.png", verDepoisButton);
                    await verDepoisButton.ClickAsync();
                    await Task.Delay(1500);
                    //await CapturarScreenshotModalAsync("modal1_ver_depois_depois.png", verDepoisButton);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"   ⚠️  Erro ao verificar botão 'Ver depois': {ex.Message}");
            }
        }

        private async Task VerificarEClicarOkEntendiAsync()
        {
            Console.WriteLine("   🔍 Verificando se existe botão 'Ok, entendi'...");

            try
            {
                await Task.Delay(Configuracao.TIMEOUT_CURTO);

                var okEntendiButton = await BuscarBotaoOkEntendiAsync();

                if (okEntendiButton != null && await okEntendiButton.IsVisibleAsync())
                {
                    Console.WriteLine("   ✅ Botão 'Ok, entendi' encontrado e visível!");
                    //await CapturarScreenshotModalAsync("modal2_ok_entendi_antes.png", okEntendiButton);
                    await okEntendiButton.ClickAsync();
                    await Task.Delay(1500);
                    //await CapturarScreenshotModalAsync("modal2_ok_entendi_depois.png", okEntendiButton);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"   ⚠️  Erro ao verificar botão 'Ok, entendi': {ex.Message}");
            }
        }

        private async Task<IElementHandle> BuscarBotaoOkEntendiAsync()
        {
            // Múltiplas estratégias de busca
            var selectors = new[]
            {
                "button.me-button.me-button-md.me-button-primary:has-text('Ok, entendi')",
                "button.me-button.me-button-md.me-button-primary",
                "button[data-cy='b-button']",
                "button:has-text('Ok, entendi')",
                "button:has-text('Ok'), button:has-text('Entendi'), button:has-text('Entendido')"
            };

            foreach (var selector in selectors)
            {
                var button = await _page.QuerySelectorAsync(selector);
                if (button != null && await button.IsVisibleAsync())
                {
                    return button;
                }
            }

            return null;
        }

        private async Task VerificarEClicarCheckboxCienteAsync()
        {
            Console.WriteLine("   🔍 Verificando se existe checkbox 'Ciente'...");

            try
            {
                await Task.Delay(Configuracao.TIMEOUT_CURTO);

                var checkbox = await BuscarCheckboxCienteAsync();

                if (checkbox != null && await checkbox.IsVisibleAsync())
                {
                    var isChecked = await checkbox.IsCheckedAsync();

                    if (!isChecked)
                    {
                        await CliqueSeguroAsync(checkbox);
                        Console.WriteLine("   ✅ Checkbox 'Ciente' clicado!");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"   ⚠️  Erro ao lidar com checkbox: {ex.Message}");
            }
        }

        private async Task<IElementHandle> BuscarCheckboxCienteAsync()
        {
            var selectors = new[]
            {
                "#ctl00_conteudo_frmComunicadoUsuario_chkCiente",
                "input[type='checkbox'][id*='chkCiente']",
                "input[type='checkbox'][name*='chkCiente']",
                "input[type='checkbox'][id*='aceite']",
                "input[type='checkbox']:has(+ label:has-text('Ciente'))"
            };

            foreach (var selector in selectors)
            {
                var element = await _page.QuerySelectorAsync(selector);
                if (element != null) return element;
            }

            return null;
        }

        private async Task VerificarEClicarBotaoGravarAsync()
        {
            Console.WriteLine("   🔍 Verificando se existe botão 'Gravar'...");

            try
            {
                await Task.Delay(1500);

                var botaoGravar = await BuscarBotaoGravarAsync();

                if (botaoGravar != null && await botaoGravar.IsVisibleAsync() && await botaoGravar.IsEnabledAsync())
                {
                    Console.WriteLine("   ✅ Botão 'Gravar' encontrado, visível e habilitado!");
                    await botaoGravar.ScrollIntoViewIfNeededAsync();
                    await Task.Delay(1000);

                    await CliqueSeguroAsync(botaoGravar);
                    Console.WriteLine("   ✅ Botão 'Gravar' clicado!");

                    await Task.Delay(4000);
                    await VerificarMensagensSistemaAsync();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"   ⚠️  Erro ao lidar com botão Gravar: {ex.Message}");
            }
        }

        private async Task<IElementHandle> BuscarBotaoGravarAsync()
        {
            var selectors = new[]
            {
                "#ctl00_conteudo_frmComunicadoUsuario_ButtonBar1_btn_ctl00_conteudo_frmComunicadoUsuario_ButtonBar1_btnAceite",
                "button[id*='btnAceite']",
                "button[onclick*='frmComunicadoUsuario']",
                "button.me-button:has-text('Gravar')",
                "button:has-text('Gravar'), button:has-text('Salvar'), button:has-text('Confirmar'), button:has-text('Aceitar')"
            };

            foreach (var selector in selectors)
            {
                var element = await _page.QuerySelectorAsync(selector);
                if (element != null) return element;
            }

            return null;
        }

        private async Task VerificarModaisRestantesAsync()
        {
            try
            {
                var modalClosed = await _page.EvaluateAsync<bool>(@"
                    () => {
                        const modals = document.querySelectorAll('.modal, .popup, [role=dialog], .dialog-overlay');
                        for (const modal of modals) {
                            const style = window.getComputedStyle(modal);
                            if (style.display !== 'none' && style.visibility !== 'hidden' && style.opacity !== '0') {
                                return false;
                            }
                        }
                        return true;
                    }
                ");

                if (!modalClosed)
                {
                    Console.WriteLine("   ⚠️  Ainda há modais abertos, tentando métodos alternativos...");
                    await _page.Keyboard.PressAsync("Escape");
                    await Task.Delay(1000);
                    await _page.ClickAsync("body", new PageClickOptions { Position = new Position { X = 10, Y = 10 } });
                }
            }
            catch { /* Ignorar erros na verificação JavaScript */ }
        }

        private async Task VerificarMensagensSistemaAsync()
        {
            try
            {
                await Task.Delay(Configuracao.TIMEOUT_CURTO);

                // Verificar mensagens de sucesso
                var msgSucesso = await _page.QuerySelectorAsync(".alert-success, .sucesso, .mensagem-sucesso, div:has-text('sucesso'), div:has-text('Sucesso!')");
                if (msgSucesso != null && await msgSucesso.IsVisibleAsync())
                {
                    var textoSucesso = await msgSucesso.InnerTextAsync();
                    Console.WriteLine($"   ✅ Mensagem de sucesso: {textoSucesso}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"   ⚠️  Erro ao verificar mensagens: {ex.Message}");
            }
        }

        private async Task CapturarScreenshotModalAsync(string nomeArquivo, IElementHandle elemento = null)
        {
            await _page.ScreenshotAsync(new PageScreenshotOptions
            {
                Path = nomeArquivo,
                FullPage = true
            });
        }

        private async Task CliqueSeguroAsync(IElementHandle elemento)
        {
            try
            {
                await elemento.ClickAsync();
            }
            catch
            {
                await _page.EvaluateAsync("arguments[0].click();", elemento);
            }
        }
    }
}