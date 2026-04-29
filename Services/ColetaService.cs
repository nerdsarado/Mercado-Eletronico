using MercadoEletronico.Models;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Playwright;
using System.ComponentModel;
using System.Globalization;
using System.Text.RegularExpressions;

namespace MercadoEletronico.Services
{
    public class ColetaService
    {
        private readonly IPage _page;
        private readonly ModalHandlerService _modalHandler;
        private int numeroLinhasTabela = 0;
        private readonly HashSet<string> _cotacoesProcessadas;
        private readonly PlanilhaControleService _planilhaControle;
        public string empresaAtual;

        public ColetaService(IPage page, ModalHandlerService modalHandler,
                             HashSet<string> cotacoesProcessadas,
                             PlanilhaControleService planilhaControle)
        {
            _page = page;
            _modalHandler = modalHandler;
            _cotacoesProcessadas = cotacoesProcessadas;
            _planilhaControle = planilhaControle;
        }

        public ColetaService(IPage page, ModalHandlerService modalHandler)
        {
            _page = page;
            _modalHandler = modalHandler;
        }

        public async Task<List<IElementHandle>> ColetarElementosDeCotacoesAsync()
        {
            Console.WriteLine("\n5. COLETANDO LISTA DE COTAÇÕES...");
            Console.WriteLine("   📋 Aguardando lista de cotações...");

            await Task.Delay(Configuracao.TIMEOUT_MEDIO);
            

            var elementos = new List<IElementHandle>();
            var elementosProcessados = new HashSet<string>();

            // **ESTRATÉGIA PRINCIPAL: Buscar pelos links das cotações baseado no HTML fornecido**
            Console.WriteLine("   🔍 Estratégia 1: Buscando por links de cotações [data-cy='open-doc-modal']...");

            // Buscar todos os links que abrem cotações
            var cotacaoLinks = await _page.QuerySelectorAllAsync("a[data-cy='open-doc-modal']");
            Console.WriteLine($"   📊 Encontrados {cotacaoLinks.Count} links de cotações");

            foreach (var link in cotacaoLinks)
            {
                try
                {
                    // Obter o texto do link (número da cotação)
                    var numeroCotacao = (await link.InnerTextAsync()).Trim();

                    if (!string.IsNullOrEmpty(numeroCotacao) && !elementosProcessados.Contains(numeroCotacao))
                    {
                        // Verificar se parece um número de cotação (apenas números)
                        if (Regex.IsMatch(numeroCotacao, @"^\d+$"))
                        {
                            elementos.Add(link);
                            elementosProcessados.Add(numeroCotacao);
                            Console.WriteLine($"   📄 Cotação encontrada: {numeroCotacao}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"   ⚠️  Erro ao verificar link de cotação: {ex.Message}");
                }
            }

            // Estratégia 2: Se não encontrou pelos links específicos, buscar por href que contenha "FornShowCotacao.asp"
            if (elementos.Count == 0)
            {
                Console.WriteLine("   🔍 Estratégia 2: Buscando por links com 'FornShowCotacao.asp'...");
                var forShowLinks = await _page.QuerySelectorAllAsync("a[href*='FornShowCotacao.asp']");
                Console.WriteLine($"   📊 Encontrados {forShowLinks.Count} links com FornShowCotacao.asp");

                foreach (var link in forShowLinks)
                {
                    try
                    {
                        var href = await link.GetAttributeAsync("href");
                        if (!string.IsNullOrEmpty(href))
                        {
                            // Extrair número da cotação do href
                            var match = Regex.Match(href, @"Cot=(\d+)");
                            if (match.Success)
                            {
                                var numeroCotacao = match.Groups[1].Value;

                                if (!elementosProcessados.Contains(numeroCotacao))
                                {
                                    elementos.Add(link);
                                    elementosProcessados.Add(numeroCotacao);
                                    Console.WriteLine($"   📄 Cotação encontrada via href: {numeroCotacao}");
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"   ⚠️  Erro ao verificar link: {ex.Message}");
                    }
                }
            }

            // Estratégia 3: Buscar por elementos que contenham números de cotação
            if (elementos.Count == 0)
            {
                Console.WriteLine("   🔍 Estratégia 3: Buscando elementos com números de cotação...");

                // Buscar todos os elementos que possam conter números de cotação
                var allElements = await _page.QuerySelectorAllAsync("td, div, span, a");

                foreach (var element in allElements)
                {
                    try
                    {
                        if (!await element.IsVisibleAsync())
                            continue;

                        var text = (await element.InnerTextAsync()).Trim();

                        // Verificar se o texto parece um número de cotação (8+ dígitos)
                        if (text.Length >= 8 && Regex.IsMatch(text, @"^\d+$"))
                        {
                            // Verificar se o elemento ou seus pais têm estrutura de cotação
                            var parentHtml = await element.EvaluateAsync<string>("el => el.parentElement.outerHTML");

                            if (parentHtml.Contains("Cotacao", StringComparison.OrdinalIgnoreCase) ||
                                parentHtml.Contains("FornShowCotacao", StringComparison.OrdinalIgnoreCase))
                            {
                                if (!elementosProcessados.Contains(text))
                                {
                                    // Tentar encontrar o link clicável
                                    var clickableElement = await element.EvaluateHandleAsync("el => el.closest('a')");
                                    if (clickableElement != null && clickableElement is IElementHandle clickableHandle)
                                    {
                                        elementos.Add(clickableHandle);
                                    }
                                    else
                                    {
                                        elementos.Add(element);
                                    }

                                    elementosProcessados.Add(text);
                                    Console.WriteLine($"   📄 Cotação encontrada pelo texto: {text}");
                                }
                            }
                        }
                    }
                    catch { /* Ignorar erros */ }
                }
            }

            Console.WriteLine($"   ✅ Total de cotações identificadas: {elementos.Count}");

            // Se encontrou cotações, imprimir debug adicional
            if (elementos.Count > 0)
            {
                Console.WriteLine("   🔍 Debug dos elementos encontrados:");
                for (int i = 0; i < Math.Min(elementos.Count, 5); i++)
                {
                    try
                    {
                        var html = await elementos[i].EvaluateAsync<string>("el => el.outerHTML");
                        Console.WriteLine($"      [{i + 1}] HTML: {html}");
                    }
                    catch { }
                }
            }

            return elementos;
        }
        public async Task<string> ExtrairNumeroCotacaoElementoAsync(IElementHandle element)
        {
            try
            {
                // Tentar extrair do texto do elemento
                var texto = await element.InnerTextAsync();
                if (!string.IsNullOrEmpty(texto) && Regex.IsMatch(texto.Trim(), @"^\d+$"))
                {
                    return texto.Trim();
                }

                // Tentar extrair do href
                var href = await element.GetAttributeAsync("href");
                if (!string.IsNullOrEmpty(href))
                {
                    var match = Regex.Match(href, @"Cot=(\d+)");
                    if (match.Success)
                    {
                        return match.Groups[1].Value;
                    }
                }

                return null;
            }
            catch
            {
                return null;
            }
        }
        private async Task<string> ExtrairNumeroDoElementoAsync(IElementHandle element)
        {
            try
            {
                // Verificar se é um link de cotação
                var tagName = await element.EvaluateAsync<string>("el => el.tagName.toLowerCase()");

                if (tagName == "a")
                {
                    // Extrair do texto do link
                    var linkText = (await element.InnerTextAsync()).Trim();
                    if (!string.IsNullOrEmpty(linkText) && Regex.IsMatch(linkText, @"^\d+$"))
                    {
                        Console.WriteLine($"      ✅ Número extraído do link: {linkText}");
                        return linkText;
                    }

                    // Extrair do href
                    var href = await element.GetAttributeAsync("href");
                    if (!string.IsNullOrEmpty(href))
                    {
                        var match = Regex.Match(href, @"Cot=(\d+)");
                        if (match.Success)
                        {
                            Console.WriteLine($"      ✅ Número extraído do href: {match.Groups[1].Value}");
                            return match.Groups[1].Value;
                        }
                    }
                }

                // Método original como fallback
                var textElement = await element.InnerTextAsync();
                var matches = Regex.Matches(textElement, @"\d+");
                if (matches.Count > 0)
                {
                    string maiorNumero = "";
                    foreach (Match match in matches)
                    {
                        if (match.Value.Length > maiorNumero.Length)
                        {
                            maiorNumero = match.Value;
                        }
                    }
                    Console.WriteLine($"      ✅ Número extraído do texto: {maiorNumero}");
                    return maiorNumero;
                }

                Console.WriteLine("      ⚠️  Nenhum número encontrado no elemento");
                return "";
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      ⚠️  Erro ao extrair número: {ex.Message}");
                return "";
            }
        }
        public async Task<List<Cotacao>> ProcessarTodasCotacoesAsync(List<IElementHandle> elementos)
        {
            var cotacoes = new List<Cotacao>();

            if (elementos.Count == 0)
            {
                Console.WriteLine("   ℹ️  Nenhuma cotação para processar.");
                return cotacoes;
            }

            for (numeroLinhasTabela = 0; numeroLinhasTabela < elementos.Count; numeroLinhasTabela++)
            {
                Console.WriteLine($"\n   📄 [{numeroLinhasTabela + 1}/{elementos.Count}] Processando cotação...");
                var cotacao = await ProcessarCotacaoClicandoAsync(elementos[numeroLinhasTabela]);

                if (cotacao != null)
                {
                    cotacoes.Add(cotacao);
                    Console.WriteLine($"      ✅ Coletada: {cotacao.NumeroCotacao}");
                }

                await VoltarParaListaAsync();
            }

            return cotacoes;
        }

        private async Task<Cotacao> ProcessarCotacaoClicandoAsync(IElementHandle cotacaoElement)
        {
            try
            {
                var numeroCotacao = await ExtrairNumeroDoElementoAsync(cotacaoElement);
                if (string.IsNullOrEmpty(numeroCotacao))
                {
                    Console.WriteLine("      ⚠️  Número da cotação não encontrado no elemento");
                    return null;
                }

                // Verificar novamente se já foi processada (redundância)
                if (_cotacoesProcessadas.Contains(numeroCotacao))
                {
                    Console.WriteLine($"      ⏭️  Cotação {numeroCotacao} já processada anteriormente");
                    return null;
                }

                Console.WriteLine($"      🔍 Cotação identificada: {numeroCotacao}");

                // **VERIFICAR SE É DA EMPRESA NESTLE**
                bool ehNestle = await VerificarSeEmpresaENestleAsync(cotacaoElement);

                if (!ehNestle)
                {
                    Console.WriteLine($"      ⏭️  Cotação {numeroCotacao} ignorada (empresa diferente de Nestle)");
                    return null;
                }

                bool temShortlist = await VerificarShortlist(cotacaoElement);
                if (temShortlist)
                {
                    Console.WriteLine($"       Shortlist identificada para a cotação {numeroCotacao} Ignorada (shortlist)");
                    return null;
                }

                // **CLICAR PARA ABRIR MODAL**
                Console.WriteLine("      🖱️  Clicando para abrir modal...");

                try
                {
                    await cotacaoElement.ClickAsync(new ElementHandleClickOptions
                    {
                        Timeout = 10000
                    });
                    Console.WriteLine("      ✅ Clique realizado!");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"      ⚠️  Erro no clique: {ex.Message}");
                    await cotacaoElement.EvaluateAsync("el => el.click()");
                    Console.WriteLine("      ✅ Clique via JavaScript");
                }

                // **AGUARDAR MODAL APARECER**
                Console.WriteLine("      ⏳ Aguardando modal aparecer...");
                await Task.Delay(3000);

                bool modalAberto = await AguardarModalIframeCarregarAsync(numeroCotacao);

                if (!modalAberto)
                {
                    Console.WriteLine("      ⚠️  Modal não apareceu completamente, mas continuando...");
                }

                //**SAIR DO COMUNICADO**
                await SairDoComunicado();

                // **AGUARDAR IFRAME CARREGAR**
                Console.WriteLine("      🌐 Aguardando iframe carregar...");
                await Task.Delay(5000);

                // **COLETAR INFORMAÇÕES DO IFRAME**
                Console.WriteLine("      📋 Coletando informações do iframe...");

                var dataVencimento = await ExtrairPrazoRespostaIframeAsync();
                var itensCotacao = await ExtrairItensCotacaoIframeAsync();
                var solicitante = await ExtrairSolicitanteIframeAsync();

                // Criar objeto de cotação
                var cotacao = new Cotacao
                {
                    NumeroCotacao = numeroCotacao,
                    Portal = "Mercado eletrônico".ToUpper(),
                    Empresa = "Nestle".ToUpper(),
                    DataVencimento = dataVencimento.Date,
                    HorarioVencimento = dataVencimento.TimeOfDay,
                    DataRegistro = DateTime.Now,
                    HorarioRegistro = DateTime.Now.TimeOfDay,
                    Solicitante = solicitante,
                    Itens = itensCotacao
                };

                Console.WriteLine($"      📅 Prazo: {dataVencimento:dd/MM/yyyy HH:mm}");
                Console.WriteLine($"      👤 Solicitante: {solicitante}");
                Console.WriteLine($"      📦 {itensCotacao.Count} itens encontrados");

                if (itensCotacao.Count > 0)
                {
                    foreach (var item in itensCotacao.Take(2))
                    {
                        Console.WriteLine($"         • {item.Substring(0, Math.Min(80, item.Length))}...");
                    }
                }
                await ImprimirCotacaoComoPDFAsync(numeroCotacao);

                // **ADICIONAR À PLANILHA DE CONTROLE**
                _planilhaControle.AdicionarCotacaoNaPlanilha(cotacao);

                // **ADICIONAR AO CONJUNTO DE PROCESSADAS**
                _cotacoesProcessadas.Add(numeroCotacao);

                // **FECHAR MODAL**
                Console.WriteLine("      ❌ Fechando modal...");
                await FecharModalComIconeXAsync();

                await Task.Delay(2000); // Aguardar modal fechar

                return cotacao;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      ❌ Erro ao processar cotação: {ex.Message}");
                Console.WriteLine($"      Stack Trace: {ex.StackTrace}");

                // Tentar fechar modal em caso de erro
                try
                {
                    await FecharModalComIconeXAsync();
                }
                catch { }

                return null;
            }
        }
        private async Task<bool> VerificarShortlist(IElementHandle cotacaoElement)
        {
            try
            {
                var buscarPaginaPrinciapl = await IdentificarShortlist(cotacaoElement);
                if (!string.IsNullOrEmpty(buscarPaginaPrinciapl) && buscarPaginaPrinciapl.Contains("ShortList", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine("      Shortlist encontrada na página principal");
                    return true;
                }
                return false;
            }
            catch(Exception ex)
            {
                Console.WriteLine("      Erro ao verificar shortlist: " + ex.Message);
                return false;
            }
        }
        private async Task<string> IdentificarShortlist(IElementHandle cotacaoElement)
        {
            try
            {
                Console.WriteLine("Verificando se há shortlist...");

                // Encontra a linha (tr) que contém o elemento da cotação
                var linha = await cotacaoElement.EvaluateHandleAsync(@"(element) => {
            // Sobe na árvore DOM até encontrar a tag TR (linha da tabela)
            let current = element;
            while (current && current.tagName !== 'TR') {
                current = current.parentElement;
            }
            return current;
        }");

                if (linha == null || linha.AsElement() == null)
                {
                    Console.WriteLine("Não foi possível encontrar a linha correspondente");
                    return "Não identificado";
                }

                var linhaElement = linha.AsElement();

                // Agora busca especificamente dentro desta linha o elemento com aria-colindex='6'
                var elementoShortlist = await linhaElement.QuerySelectorAsync("[aria-colindex='6']");

                if (elementoShortlist != null)
                {
                    var texto = await elementoShortlist.InnerTextAsync();
                    texto = texto.Trim();

                    if (!string.IsNullOrEmpty(texto))
                    {
                        Console.WriteLine($"      🏢 Elemento encontrado na linha atual (ShortList): {texto}");

                        if (texto.Contains("ShortList", StringComparison.OrdinalIgnoreCase))
                        {
                            Console.WriteLine("Shortlist encontrada!");
                            return texto;
                        }
                        else
                        {
                            Console.WriteLine("Elemento encontrado, mas não é Shortlist");
                        }
                    }
                }
                else
                {
                    Console.WriteLine("Nenhum elemento com aria-colindex='6' encontrado nesta linha");
                }

                return "Não identificado";
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro ao identificar shortlist: {ex.Message}");
                return "Não identificado";
            }
        }

        private async Task<bool> AguardarModalIframeCarregarAsync(string numeroCotacao)
        {
            var timeout = 15000; // 15 segundos
            var startTime = DateTime.Now;

            Console.WriteLine("      ⏳ Aguardando modal com iframe...");

            while ((DateTime.Now - startTime).TotalMilliseconds < timeout)
            {
                try
                {
                    // Verificar se o modal apareceu
                    var modal = await _page.QuerySelectorAsync(".modal-dialog.modal-xl, [class*='modal'], .modal-content");

                    if (modal != null && await modal.IsVisibleAsync())
                    {
                        Console.WriteLine("      ✅ Modal detectado!");

                        // Verificar se tem iframe
                        var iframe = await _page.QuerySelectorAsync("#document-details-iframe, iframe.frame-in-modal, iframe");

                        if (iframe != null)
                        {
                            Console.WriteLine("      ✅ Iframe detectado!");

                            // Verificar se o iframe tem o src correto
                            var src = await iframe.GetAttributeAsync("src");
                            if (!string.IsNullOrEmpty(src) && src.Contains(numeroCotacao))
                            {
                                Console.WriteLine($"      ✅ Iframe com cotação {numeroCotacao}!");
                                return true;
                            }
                        }

                        // Verificar se há o número da cotação no breadcrumb
                        var breadcrumbText = await modal.InnerTextAsync();
                        if (breadcrumbText.Contains(numeroCotacao))
                        {
                            Console.WriteLine($"      ✅ Número {numeroCotacao} no modal!");
                            return true;
                        }
                    }

                    // Verificar por elementos específicos do header do modal
                    var header = await _page.QuerySelectorAsync(".modal-header, .doc-modal__header");
                    if (header != null && await header.IsVisibleAsync())
                    {
                        Console.WriteLine("      ✅ Header do modal detectado!");
                        return true;
                    }

                    await Task.Delay(1000);
                    Console.Write(".");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"\n      ⚠️  Erro na verificação: {ex.Message}");
                }
            }

            Console.WriteLine($"\n      ⚠️  Timeout após {timeout / 1000} segundos");
            return false;
        }
        private async Task SairDoComunicado()
        {
            try
            {
                Console.WriteLine("  Verificando se possui comunicado...");
                var frame = await _page.QuerySelectorAsync("#document-details-iframe, iframe.frame-in-modal");
                if(frame != null)
                {
                    var frameElement = await frame.ContentFrameAsync();

                    if (frameElement != null)
                    {
                        var comunicado = await frameElement.QuerySelectorAsync(".headerTituloPaginaME_text");
                        if (comunicado != null)
                        {

                            // Dar tempo para o iframe carregar
                            await frameElement.WaitForLoadStateAsync(LoadState.DOMContentLoaded);
                            await Task.Delay(5000);

                            // Buscar o elemento de prazo DENTRO do iframe
                            Console.WriteLine("      📅 Procurando Input...");
                            var comunicadoElement = await frameElement.QuerySelectorAsync("#ctl00_conteudo_frmComunicadoUsuario_chkCiente");
                            if (comunicadoElement != null)
                            {
                                Console.WriteLine("      📅 Possui comunicado, clicando em 'Li e estou ciente'...");
                                await comunicadoElement.ClickAsync();

                            }
                            else
                            {
                                Console.WriteLine("      ⚠️  Elemento de comunicado não encontrado no iframe");

                            }
                            Console.WriteLine("      📅 Procurando Botão Gravar...");
                            var botaoGravar = await frameElement.QuerySelectorAsync("#ctl00_conteudo_frmComunicadoUsuario_ButtonBar1_btn_ctl00_conteudo_frmComunicadoUsuario_ButtonBar1_btnAceite");
                            if (botaoGravar != null)
                            {
                                Console.WriteLine("      ⚠️  Botão Gravar encontrado no Iframe, clicando no botão...");
                                await botaoGravar.ClickAsync();
                            }
                        }
                        else
                        {
                            Console.WriteLine("Não possui comunicado, continuando...");
                        }
                    }
                    else
                    {
                        Console.WriteLine("      ⚠️  Não conseguiu acessar conteúdo do iframe");
                    }
                   
                }
                else
                {
                    Console.WriteLine("Não foi possivel continuar com a tarefa, pois o modal não foi identificado.");
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine($"      ⚠️  Erro ao sair do comunicado: {ex.Message}");
            }
        }

        private async Task<DateTime> ExtrairPrazoRespostaIframeAsync()
        {
            try
            {
                Console.WriteLine("      📅 Extraindo prazo do iframe...");

                // **MUDANÇA IMPORTANTE: Acessar o conteúdo do iframe**
                var frame = await _page.QuerySelectorAsync("#document-details-iframe, iframe.frame-in-modal");

                if (frame == null)
                {
                    Console.WriteLine("      ⚠️  Iframe não encontrado, tentando na página principal");
                    return await ExtrairPrazoRespostaAsync(); // Método original
                }

                // Obter o Frame do iframe
                var frameElement = await frame.ContentFrameAsync();

                if (frameElement == null)
                {
                    Console.WriteLine("      ⚠️  Não conseguiu acessar conteúdo do iframe");
                    return DateTime.Now.AddDays(7);
                }

                // Dar tempo para o iframe carregar
                await frameElement.WaitForLoadStateAsync(LoadState.DOMContentLoaded);
                await Task.Delay(2000);

                // Buscar o elemento de prazo DENTRO do iframe
                var prazoElement = await frameElement.QuerySelectorAsync("#tdDadosCotacaoPrazoResposta");

                if (prazoElement != null)
                {
                    var textoPrazo = await prazoElement.InnerTextAsync();
                    Console.WriteLine($"      📅 Texto do prazo: '{textoPrazo}'");

                    textoPrazo = textoPrazo.Replace("&nbsp;", " ").Trim();

                    // Tentar múltiplos formatos
                    var formatos = new[] {
                "dd/MM/yyyy HH:mm",
                "dd/MM/yyyy HH:mm:ss",
                "dd/MM/yyyy"
            };

                    foreach (var formato in formatos)
                    {
                        if (DateTime.TryParseExact(textoPrazo, formato,
                            CultureInfo.InvariantCulture, DateTimeStyles.None, out var data))
                        {
                            Console.WriteLine($"      ✅ Data extraída: {data:dd/MM/yyyy HH:mm}");
                            return data;
                        }
                    }

                    // Tentar regex
                    var match = Regex.Match(textoPrazo, @"(\d{2}/\d{2}/\d{4})\s*(\d{2}:\d{2})");
                    if (match.Success)
                    {
                        var dataStr = match.Groups[1].Value + " " + match.Groups[2].Value;
                        if (DateTime.TryParseExact(dataStr, "dd/MM/yyyy HH:mm",
                            CultureInfo.InvariantCulture, DateTimeStyles.None, out var data))
                        {
                            Console.WriteLine($"      ✅ Data extraída via regex: {data:dd/MM/yyyy HH:mm}");
                            return data;
                        }
                    }
                }
                else
                {
                    Console.WriteLine("      ⚠️  Elemento #tdDadosCotacaoPrazoResposta não encontrado no iframe");

                    // Tentar outras estratégias dentro do iframe
                    var todoConteudo = await frameElement.ContentAsync();
                    if (todoConteudo.Contains("Prazo") || todoConteudo.Contains("Vencimento"))
                    {
                        Console.WriteLine("      🔍 Texto 'Prazo' ou 'Vencimento' encontrado no iframe");

                        // Procurar datas no conteúdo todo
                        var regexData = new Regex(@"(\d{2}/\d{2}/\d{4}\s*\d{2}:\d{2})");
                        var match = regexData.Match(todoConteudo);
                        if (match.Success)
                        {
                            var dataStr = match.Value.Replace(" ", "");
                            if (DateTime.TryParseExact(dataStr, "dd/MM/yyyyHH:mm",
                                CultureInfo.InvariantCulture, DateTimeStyles.None, out var data))
                            {
                                Console.WriteLine($"      ✅ Data encontrada no conteúdo: {data:dd/MM/yyyy HH:mm}");
                                return data;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      ⚠️  Erro ao extrair prazo do iframe: {ex.Message}");
            }

            Console.WriteLine("      ⚠️  Usando data padrão (7 dias)");
            return DateTime.Now.AddDays(7);
        }

        private async Task<List<string>> ExtrairItensCotacaoIframeAsync()
        {
            var itens = new List<string>();

            try
            {
                Console.WriteLine("      📦 Extraindo itens do iframe...");

                var frame = await _page.QuerySelectorAsync("#document-details-iframe, iframe.frame-in-modal");

                if (frame == null)
                {
                    Console.WriteLine("      ⚠️  Iframe não encontrado");
                    return itens;
                }

                var frameElement = await frame.ContentFrameAsync();

                if (frameElement == null)
                {
                    Console.WriteLine("      ⚠️  Não conseguiu acessar conteúdo do iframe");
                    return itens;
                }

                // Buscar elementos específicos DENTRO do iframe
                var itemElements = await frameElement.QuerySelectorAllAsync("span[id^='spanContratoID']");
                Console.WriteLine($"      📦 Encontrados {itemElements.Count} elementos de itens no iframe");

                foreach (var itemElement in itemElements)
                {
                    try
                    {
                        var textoItem = await itemElement.InnerTextAsync();
                        if (!string.IsNullOrWhiteSpace(textoItem))
                        {
                            textoItem = Regex.Replace(textoItem, @"\s+", " ").Trim();

                            // Remove o código entre parênteses (ex: "Nome do Item (12345)" -> "Nome do Item")
                            textoItem = Regex.Replace(textoItem, @"\s*\([^)]*\)", "").Trim();

                            itens.Add(textoItem);
                            Console.WriteLine($"      📦 Item: {textoItem.Substring(0, Math.Min(80, textoItem.Length))}...");
                        }
                    }
                    catch { }
                }

                // Se não encontrou, buscar por "Código Produto" dentro do iframe
                if (itens.Count == 0)
                {
                    var possiveisItens = await frameElement.QuerySelectorAllAsync("*");

                    foreach (var elemento in possiveisItens.Take(50)) // Limitar busca
                    {
                        try
                        {
                            var texto = await elemento.InnerTextAsync();
                            if (texto.Contains("Código Produto") && texto.Length > 20)
                            {
                                texto = Regex.Replace(texto, @"\s+", " ").Trim();

                                // Remove o código entre parênteses
                                texto = Regex.Replace(texto, @"\s*\([^)]*\)", "").Trim();

                                itens.Add(texto);
                                Console.WriteLine($"      📦 Item via texto: {texto.Substring(0, Math.Min(80, texto.Length))}...");
                            }
                        }
                        catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      ⚠️  Erro ao extrair itens do iframe: {ex.Message}");
            }

            return itens;
        }

        private async Task FecharModalComIconeXAsync()
        {
            try
            {
                Console.WriteLine("      ❌ Procurando botão X para fechar modal...");

                // Procurar pelo ícone X específico
                var botaoX = await _page.QuerySelectorAsync(".icon-xmark, i.icon-xmark, .me-icon-l.icon-xmark");

                if (botaoX != null)
                {
                    Console.WriteLine("      ✅ Ícone X encontrado!");

                    // Subir na hierarquia para encontrar o botão clicável
                    var botaoPai = await botaoX.EvaluateHandleAsync(@"
                el => {
                    // Encontrar o botão pai
                    let current = el.parentElement;
                    while (current) {
                        if (current.tagName === 'BUTTON' || 
                            current.hasAttribute('onclick') ||
                            current.classList.contains('btn') ||
                            current.classList.contains('button')) {
                            return current;
                        }
                        current = current.parentElement;
                    }
                    return el.parentElement || el;
                }
            ");

                    if (botaoPai is IElementHandle elementoParaClicar)
                    {
                        await elementoParaClicar.ClickAsync();
                        Console.WriteLine("      ✅ Modal fechado com botão X!");
                        return;
                    }
                }

                // Se não encontrou o ícone X específico, tentar outras estratégias
                Console.WriteLine("      🔍 Tentando outros métodos para fechar modal...");

                // Botão com classe de fechar
                var botoesFechar = await _page.QuerySelectorAllAsync("button.close, [data-dismiss=modal], [aria-label=Close]");

                foreach (var botao in botoesFechar)
                {
                    try
                    {
                        if (await botao.IsVisibleAsync())
                        {
                            await botao.ClickAsync();
                            Console.WriteLine("      ✅ Modal fechado com botão de fechar");
                            return;
                        }
                    }
                    catch { }
                }

                // Clicar fora do modal
                await _page.ClickAsync("body", new PageClickOptions
                {
                    Position = new Position { X = 10, Y = 10 }
                });
                Console.WriteLine("      ✅ Clicado fora do modal");

                // Tecla ESC
                await _page.Keyboard.PressAsync("Escape");
                Console.WriteLine("      ✅ Pressionada tecla ESC");

                await Task.Delay(1000);

            }
            catch (Exception ex)
            {
                Console.WriteLine($"      ⚠️  Erro ao fechar modal: {ex.Message}");
            }
        }

        // Adicionar métodos similares para extrair solicitante, vendedor e status do iframe
        private async Task<string> ExtrairSolicitanteIframeAsync()
        {
            try
            {
                var frame = await _page.QuerySelectorAsync("#document-details-iframe, iframe.frame-in-modal");
                if (frame == null) return "Não identificado";

                var frameElement = await frame.ContentFrameAsync();
                if (frameElement == null) return "Não identificado";

                // Buscar por "Solicitante", "Cliente", etc no iframe
                var solicitanteElement = await frameElement.QuerySelectorAsync("#tdDadosCotacaoSolicitante");
                if(solicitanteElement != null)
                {
                    var textoSolicitante = await solicitanteElement.InnerTextAsync();
                    if(!string.IsNullOrEmpty(textoSolicitante))
                    {
                        Console.WriteLine($"      📋 Solicitante extraído do iframe: {textoSolicitante}");
                        return textoSolicitante;
                    }
                }
                else
                {
                    Console.WriteLine("      ⚠️  Elemento #tdDadosCotacaoSolicitante não encontrado no iframe.");
                }
            }
            catch { }

            return "Não identificado";
        }
        private async Task<DateTime> ExtrairPrazoRespostaAsync()
        {
            try
            {
                Console.WriteLine("      📅 Extraindo prazo para resposta...");

                // Primeiro, dar um pequeno delay extra para garantir
                await Task.Delay(500);

                // Procurar pelo elemento específico
                var prazoElement = await _page.QuerySelectorAsync("#tdDadosCotacaoPrazoResposta");

                if (prazoElement != null)
                {
                    // Verificar visibilidade
                    if (!await prazoElement.IsVisibleAsync())
                    {
                        Console.WriteLine("      ⚠️  Elemento de prazo encontrado mas não visível");
                        await prazoElement.ScrollIntoViewIfNeededAsync();
                        await Task.Delay(500);
                    }

                    var textoPrazo = await prazoElement.InnerTextAsync();
                    Console.WriteLine($"      📅 Texto do prazo: '{textoPrazo}'");

                    // Limpar o texto
                    textoPrazo = textoPrazo.Replace("&nbsp;", " ").Trim();

                    // Tentar múltiplos formatos de data
                    var formatos = new[] {
                "dd/MM/yyyy HH:mm",
                "dd/MM/yyyy HH:mm:ss",
                "dd/MM/yyyy",
                "d/MM/yyyy HH:mm",
                "d/M/yyyy HH:mm"
            };

                    foreach (var formato in formatos)
                    {
                        if (DateTime.TryParseExact(textoPrazo, formato,
                            CultureInfo.InvariantCulture, DateTimeStyles.None, out var data))
                        {
                            Console.WriteLine($"      ✅ Data de vencimento extraída: {data:dd/MM/yyyy HH:mm}");
                            return data;
                        }
                    }

                    // Tentar extrair com regex
                    var match = Regex.Match(textoPrazo, @"(\d{1,2}/\d{1,2}/\d{4})\s*(\d{1,2}:\d{2})");
                    if (match.Success)
                    {
                        var dataStr = match.Groups[1].Value + " " + match.Groups[2].Value;
                        if (DateTime.TryParseExact(dataStr, "dd/MM/yyyy HH:mm",
                            CultureInfo.InvariantCulture, DateTimeStyles.None, out var data))
                        {
                            Console.WriteLine($"      ✅ Data de vencimento extraída via regex: {data:dd/MM/yyyy HH:mm}");
                            return data;
                        }
                    }
                }
                else
                {
                    Console.WriteLine("      ⚠️  Elemento #tdDadosCotacaoPrazoResposta não encontrado");

                    // Tentar encontrar por texto
                    var elementosComPrazo = await _page.QuerySelectorAllAsync("*");
                    foreach (var elemento in elementosComPrazo)
                    {
                        try
                        {
                            if (await elemento.IsVisibleAsync())
                            {
                                var texto = await elemento.InnerTextAsync();
                                if (texto.Contains("/2025") || texto.Contains("/2024") || texto.Contains("/2026"))
                                {
                                    Console.WriteLine($"      🔍 Possível data encontrada: '{texto.Substring(0, Math.Min(50, texto.Length))}'");
                                }
                            }
                        }
                        catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      ⚠️  Erro ao extrair prazo: {ex.Message}");
            }

            Console.WriteLine("      ⚠️  Usando data padrão (7 dias)");
            return DateTime.Now.AddDays(7);
        }
        private async Task FecharModalAsync()
        {
            try
            {
                // Tentar múltiplas estratégias para fechar o modal
                var fechou = false;

                // Estratégia 1: Botão de fechar (X)
                var botoesFechar = await _page.QuerySelectorAllAsync(".close, [data-dismiss=modal], [aria-label=Close], button:has-text('×')");
                foreach (var botao in botoesFechar)
                {
                    try
                    {
                        if (await botao.IsVisibleAsync())
                        {
                            await botao.ClickAsync();
                            Console.WriteLine("      ✅ Modal fechado com botão X");
                            fechou = true;
                            break;
                        }
                    }
                    catch { }
                }

                // Estratégia 2: Botão "Fechar", "Cancelar", "Voltar"
                if (!fechou)
                {
                    var botoesTexto = await _page.QuerySelectorAllAsync("button:has-text('Fechar'), button:has-text('Cancelar'), button:has-text('Voltar')");
                    foreach (var botao in botoesTexto)
                    {
                        try
                        {
                            if (await botao.IsVisibleAsync())
                            {
                                await botao.ClickAsync();
                                Console.WriteLine("      ✅ Modal fechado com botão de texto");
                                fechou = true;
                                break;
                            }
                        }
                        catch { }
                    }
                }

                // Estratégia 3: Clicar fora do modal
                if (!fechou)
                {
                    await _page.ClickAsync("body", new PageClickOptions
                    {
                        Position = new Position { X = 10, Y = 10 }
                    });
                    Console.WriteLine("      ✅ Clicado fora do modal");
                    fechou = true;
                }

                // Estratégia 4: Tecla ESC
                if (!fechou)
                {
                    await _page.Keyboard.PressAsync("Escape");
                    Console.WriteLine("      ✅ Pressionada tecla ESC");
                }

                await Task.Delay(1000); // Aguardar fechamento
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      ⚠️  Erro ao fechar modal: {ex.Message}");
            }
        }
        private async Task<bool> VerificarSeEstaEmModalAsync()
        {
            try
            {
                // Verificar se há elementos de modal visíveis
                var modalVisivel = await _page.EvaluateAsync<bool>(@"
            () => {
                const modals = document.querySelectorAll('.modal, .modal-dialog, [role=dialog]');
                for (const modal of modals) {
                    const style = window.getComputedStyle(modal);
                    if (style.display !== 'none' && style.visibility !== 'hidden' && style.opacity !== '0') {
                        return true;
                    }
                }
                return false;
            }
        ");

                return modalVisivel;
            }
            catch
            {
                return false;
            }
        }
        private async Task VoltarParaListaAsync()
        {
            try
            {
                Console.WriteLine("      ↩️  Voltando para lista de cotações...");

                // Primeiro verificar se há modal aberto
                if (await VerificarSeEstaEmModalAsync())
                {
                    Console.WriteLine("      ⚠️  Modal aberto detectado, fechando primeiro...");
                    await FecharModalAsync();
                    await Task.Delay(2000);
                }

                // Verificar se já está na lista
                var content = await _page.ContentAsync();
                if (content.Contains("FornShowCotacao.asp", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine("      ✅ Já está na lista de cotações");
                    return;
                }

                // Tentar voltar
                await _page.GoBackAsync();
                await _page.WaitForLoadStateAsync(LoadState.DOMContentLoaded);
                await Task.Delay(3000);

                Console.WriteLine("      ✅ Voltou para lista de cotações");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      ⚠️  Erro ao voltar: {ex.Message}");

                // Tentar recarregar a página
                try
                {
                    await _page.GotoAsync(_page.Url);
                    await Task.Delay(3000);
                    Console.WriteLine("      ✅ Página recarregada");
                }
                catch { }
            }
        }
        private async Task<bool> VerificarSeEmpresaENestleAsync(IElementHandle cotacaoElement)
        {
            try
            {
                Console.WriteLine("      🏢 Verificando empresa solicitante...");
                var buscarPaginaPrinciapl = await BuscarEmpresaNaPaginaPrincipalAsync(cotacaoElement);
                if(!string.IsNullOrEmpty(buscarPaginaPrinciapl) && buscarPaginaPrinciapl.Contains("Nestle", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine("      ✅ Empresa Nestle encontrada na página principal");
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      ⚠️  Erro ao verificar empresa: {ex.Message}");
                // Em caso de erro, assumir que não é Nestle para segurança
                return false;
            }
        }
        private async Task<string> ColetarInformacoesBasicasAsync(IElementHandle cotacaoElement)
        {
            try
            {
                // Tentar obter informações do elemento pai ou adjacente
                var html = await cotacaoElement.EvaluateAsync<string>("el => el.outerHTML");
                var texto = await cotacaoElement.InnerTextAsync();

                // Verificar se contém "Nestle" no HTML ou texto
                if (html.Contains("Nestle", StringComparison.OrdinalIgnoreCase) ||
                    texto.Contains("Nestle", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine("      ✅ Empresa Nestle encontrada diretamente no elemento da cotação");
                    return "Nestle";
                }

                // Buscar elemento pai que possa conter a empresa
                var elementoPai = await cotacaoElement.EvaluateHandleAsync(@"el => {
                // Subir na hierarquia para encontrar elemento com classe específica
                let current = el.parentElement;
                for (let i = 0; i < 5 && current; i++) {
                    if (current.classList && 
                        (current.classList.contains('left-info') || 
                         current.classList.contains('doc-title') ||
                         current.querySelector('[data-cy=\""cell-content-value\""]'))) {
                        return current;
                    }
                    current = current.parentElement;
                }
                return el.parentElement;
                }");
        
               if (elementoPai != null && elementoPai is IElementHandle pai)
               {
                 var textoPai = await pai.InnerTextAsync();
                 if (textoPai.Contains("Nestle", StringComparison.OrdinalIgnoreCase))
                 {
                        Console.WriteLine("      ✅ Empresa Nestle encontrada no elemento pai da cotação");
                        return "Nestle";
                 }

                 // Buscar elemento específico com data-cy="cell-content-value"
                 var empresaElement = await pai.QuerySelectorAsync("[data-cy='cell-content-value']");
                 if (empresaElement != null)
                 {
                 var empresaTexto = await empresaElement.InnerTextAsync();
                 empresaTexto = empresaTexto.Trim();
                 Console.WriteLine($"      🏢 Empresa encontrada nas infos básicas: {empresaTexto}");
                 return empresaTexto;
                 }
               }
            }
            catch (Exception ex)
            {
             Console.WriteLine($"      ⚠️  Erro ao coletar infos básicas: {ex.Message}");
            }

            return string.Empty;
        }
        private async Task<string> BuscarEmpresaNaPaginaPrincipalAsync(IElementHandle cotacaoElement)
        {
            try
            {
                Console.WriteLine("Verificando se há Nestle...");

                // Encontra a linha (tr) que contém o elemento da cotação
                var linha = await cotacaoElement.EvaluateHandleAsync(@"(element) => {
            // Sobe na árvore DOM até encontrar a tag TR (linha da tabela)
            let current = element;
            while (current && current.tagName !== 'TR') {
                current = current.parentElement;
            }
            return current;
        }");

                if (linha == null || linha.AsElement() == null)
                {
                    Console.WriteLine("Não foi possível encontrar a linha correspondente");
                    return "Não identificado";
                }

                var linhaElement = linha.AsElement();

                // Agora busca especificamente dentro desta linha o elemento com aria-colindex='2'
                var elementoShortlist = await linhaElement.QuerySelectorAsync("[aria-colindex='2']");

                if (elementoShortlist != null)
                {
                    var texto = await elementoShortlist.InnerTextAsync();
                    texto = texto.Trim();

                    if (!string.IsNullOrEmpty(texto))
                    {
                        Console.WriteLine($"      🏢 Elemento encontrado na linha atual (Nestle): {texto}");

                        if (texto == "NESTLE")
                        {
                            Console.WriteLine("Nestle encontrada!");
                            return texto;
                        }
                        else
                        {
                            Console.WriteLine("Elemento encontrado, mas não é Nestle");
                        }
                    }
                }
                else
                {
                    Console.WriteLine("Nenhum elemento com aria-colindex='2' encontrado nesta linha");
                }

                return "Não identificado";
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro ao identificar Nestle: {ex.Message}");
                return "Não identificado";
            }
        }
        private async Task ImprimirCotacaoComoPDFAsync(string numeroCotacao)
        {
            try
            {
                var frame = await _page.QuerySelectorAsync("#document-details-iframe, iframe.frame-in-modal");
                if (frame == null) Console.WriteLine("Não identificado");

                var frameElement = await frame.ContentFrameAsync();
                if (frameElement == null) Console.WriteLine("Não identificado");

                Console.WriteLine($"      🖨️  Iniciando impressão da cotação {numeroCotacao} como PDF...");

                // Aguardar a abertura da nova janela
                var novaPaginaTask = _page.Context.WaitForPageAsync();

                // Clicar no botão de imprimir (ajuste o seletor conforme necessário)
                var botaoImprimir = frameElement.Locator("#MEComponentManager_MEButton_1").First;
                if (botaoImprimir == null)
                {
                    Console.WriteLine("      ⚠️  Botão de imprimir não encontrado");
                    return;
                }

                // Clicar no botão que abre a nova janela
                await botaoImprimir.ClickAsync();

                // Aguardar a nova página abrir
                var novaPagina = await novaPaginaTask;

                Console.WriteLine("      ✅ Nova janela de impressão detectada");

                // Aguardar a página de impressão carregar
                await novaPagina.WaitForLoadStateAsync(LoadState.DOMContentLoaded);
                await Task.Delay(2000); // Delay extra para garantir carregamento

                // Configurar o caminho do PDF
                string pastaDestino = @"\\SERVIDOR2\Publico\ALLAN\MERCADO-ELETRONICO";
                string caminhoPDF = Path.Combine(pastaDestino, $"Cotacao_{numeroCotacao}.pdf");

                // Garantir que a pasta existe
                Directory.CreateDirectory(pastaDestino);

                // Gerar PDF diretamente (se a página tiver o conteúdo)
                await novaPagina.PdfAsync(new PagePdfOptions
                {
                    Path = caminhoPDF,
                    Format = "A4",
                    PrintBackground = true,
                    Margin = new Margin
                    {
                        Top = "20px",
                        Bottom = "20px",
                        Left = "20px",
                        Right = "20px"
                    }
                });

                Console.WriteLine($"      ✅ PDF salvo em: {caminhoPDF}");

                // Fechar a nova janela
                await novaPagina.CloseAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      ⚠️  Erro ao imprimir PDF: {ex.Message}");
            }
        }
        public async Task Executar(IPage page, ModalHandlerService modalHandler, PlanilhaControleService planilhaControle, HashSet<string> cotacoesProcessadas, string conta)
        {
            try
            {
                var contas = new Models.Contas();
         
                        // 3. Serviços
                        var loginService = new LoginService(page);
                        var navegacaoService = new NavegacaoService(page);
                        var exportadorService = new ExportadorService();
                        modalHandler = new ModalHandlerService(page);
                        ColetaService coletaService = new ColetaService(page, modalHandler, cotacoesProcessadas, planilhaControle);

                        await loginService.RealizarLoginAsync(conta);
                        await modalHandler.LidarComTodosModaisAsync();

                        // Verificar se já está na página correta antes de navegar
                        Console.WriteLine($"\n📍 URL após login: {page.Url}");

                        await navegacaoService.NavegarParaOportunidadesAsync();
                        Console.WriteLine("\n🔍 VERIFICANDO CONTEÚDO DA PÁGINA...");
                        var pageContent = await page.ContentAsync();
                        Console.WriteLine($"   Tamanho do conteúdo: {pageContent.Length} caracteres");

                        // Verificar se há cotações no conteúdo
                        if (pageContent.Contains("FornShowCotacao.asp", StringComparison.OrdinalIgnoreCase))
                        {
                            Console.WriteLine("   ✅ Links de cotações encontrados no HTML");
                        }

                        // Procurar por números no conteúdo que pareçam cotações
                        var cotacaoMatches = Regex.Matches(pageContent, @"Cot=(\d+)");
                        Console.WriteLine($"   Encontrados {cotacaoMatches.Count} números de cotação no HTML");

                        var elementos = await coletaService.ColetarElementosDeCotacoesAsync();

                        // Filtrar elementos já processados
                        var elementosNaoProcessados = new List<IElementHandle>();
                        foreach (var elemento in elementos)
                        {
                            try
                            {
                                var numero = await coletaService.ExtrairNumeroCotacaoElementoAsync(elemento);
                                if (!string.IsNullOrEmpty(numero) && !cotacoesProcessadas.Contains(numero))
                                {
                                    elementosNaoProcessados.Add(elemento);
                                }
                                else if (!string.IsNullOrEmpty(numero))
                                {
                                    Console.WriteLine($"   ⏭️  Cotação {numero} já processada, pulando...");
                                }
                            }
                            catch { }
                        }

                        Console.WriteLine($"\n📋 Encontradas {elementos.Count} cotações no total");
                        Console.WriteLine($"📋 Cotações não processadas: {elementosNaoProcessados.Count}");

                        if (elementosNaoProcessados.Count > 0)
                        {
                            var cotacoes = await coletaService.ProcessarTodasCotacoesAsync(elementosNaoProcessados);
                            Console.WriteLine($"\n7. EXPORTANDO PARA EXCEL...");
                            exportadorService.ExportarParaExcel(cotacoes);
                            Console.WriteLine($"\n✅ Processo concluído! {cotacoes.Count} novas cotações coletadas.");
                        }
                        else
                        {
                            Console.WriteLine($"\nℹ️  Nenhuma cotação nova encontrada para processar.");
                        }
                
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\n❌ ERRO: {ex.Message}");
                Console.WriteLine($"Stack Trace: {ex.StackTrace}");
            }
            finally
            {
                if (page != null)
                {
                    await page.CloseAsync();
                }
            }
        }
    }
}