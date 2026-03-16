using MercadoEletronico.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Text.RegularExpressions;

namespace MercadoEletronico.Services
{
    public class PlanilhaControleService
    {
        private readonly string _caminhoPlanilha;
        private FileSystemWatcher _fileWatcher;
        private readonly object _lock = new object();

        // Constantes para as colunas
        private const int COL_NUMERO_COTACAO = 1;
        private const int COL_PORTAL = 2;
        private const int COL_CLIENTE = 3;
        private const int COL_DATA_VENCIMENTO = 4;
        private const int COL_HORARIO_VENCIMENTO = 5;
        private const int COL_PRODUTO = 6;
        private const int COL_DATA_ENTREGA = 7;
        private const int COL_HORARIO_ENTREGA = 8;
        private const int COL_EMPRESA = 9;

        public PlanilhaControleService()
        {
            // Caminho da planilha
            _caminhoPlanilha = @"\\SERVIDOR2\Publico\ANMYNA\CONTROLE PEDIDOS EQUIPE ANMYNA 2025 (Salvo automaticamente).xlsx";

            // Configurar permissões do EPPlus
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Inicializar watcher para detectar alterações na planilha
            InicializarFileWatcher();
        }

        private void InicializarFileWatcher()
        {
            try
            {
                string diretorio = Path.GetDirectoryName(_caminhoPlanilha);
                string arquivo = Path.GetFileName(_caminhoPlanilha);

                _fileWatcher = new FileSystemWatcher(diretorio, arquivo)
                {
                    NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.Size
                };

                _fileWatcher.Changed += OnPlanilhaAlterada;
                _fileWatcher.EnableRaisingEvents = true;

                Console.WriteLine("   👀 Monitorando alterações na planilha de controle...");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"   ⚠️  Erro ao iniciar monitoramento da planilha: {ex.Message}");
            }
        }

        private void OnPlanilhaAlterada(object sender, FileSystemEventArgs e)
        {
            // Pequeno delay para garantir que o arquivo foi totalmente salvo
            System.Threading.Thread.Sleep(1000);
            Console.WriteLine($"   📊 Planilha de controle foi atualizada em {DateTime.Now:HH:mm:ss}");
        }

        public HashSet<string> CarregarCotacoesProcessadas()
        {
            var cotacoesProcessadas = new HashSet<string>();

            try
            {
                Console.WriteLine($"   📊 Carregando cotações já processadas da planilha...");

                if (!File.Exists(_caminhoPlanilha))
                {
                    Console.WriteLine($"   ⚠️  Planilha não encontrada: {_caminhoPlanilha}");
                    Console.WriteLine($"   🔍 Criando nova planilha de controle...");
                    CriarPlanilhaInicial();
                    return cotacoesProcessadas;
                }

                lock (_lock)
                {
                    using (var package = new ExcelPackage(new FileInfo(_caminhoPlanilha)))
                    {
                        // Usar a primeira planilha
                        var worksheet = package.Workbook.Worksheets[0];

                        if (worksheet == null)
                        {
                            Console.WriteLine("   ⚠️  Nenhuma planilha encontrada no arquivo");
                            return cotacoesProcessadas;
                        }

                        // Verificar se tem cabeçalho na primeira linha
                        if (worksheet.Cells[1, COL_NUMERO_COTACAO].Value == null)
                        {
                            Console.WriteLine("   ⚠️  Planilha sem cabeçalho, criando cabeçalhos...");
                            CriarCabecalhos(worksheet);
                            package.Save();
                            return cotacoesProcessadas;
                        }

                        int linhaAtual = 2; // Linha 1 é cabeçalho
                        while (linhaAtual <= 10000) // Limite de segurança
                        {
                            var valor = worksheet.Cells[linhaAtual, COL_NUMERO_COTACAO].Value?.ToString();

                            if (string.IsNullOrEmpty(valor))
                                break;

                            // Extrair apenas números se necessário
                            var numeroCotacao = Regex.Match(valor, @"\d+").Value;

                            if (!string.IsNullOrEmpty(numeroCotacao))
                            {
                                cotacoesProcessadas.Add(numeroCotacao);
                            }

                            linhaAtual++;
                        }

                        Console.WriteLine($"   ✅ Carregadas {cotacoesProcessadas.Count} cotações já processadas");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"   ⚠️  Erro ao carregar cotações processadas: {ex.Message}");
            }

            return cotacoesProcessadas;
        }

        private void CriarCabecalhos(ExcelWorksheet worksheet)
        {
            // Criar cabeçalhos na ordem especificada
            worksheet.Cells[1, COL_NUMERO_COTACAO].Value = "NÚMERO DA COTAÇÃO";
            worksheet.Cells[1, COL_PORTAL].Value = "PORTAL";
            worksheet.Cells[1, COL_CLIENTE].Value = "CLIENTE";
            worksheet.Cells[1, COL_DATA_VENCIMENTO].Value = "DATA DE VENCIMENTO";
            worksheet.Cells[1, COL_HORARIO_VENCIMENTO].Value = "HORÁRIO DE VENCIMENTO";
            worksheet.Cells[1, COL_PRODUTO].Value = "PRODUTO";
            worksheet.Cells[1, COL_DATA_ENTREGA].Value = "DATA DE ENTREGA";
            worksheet.Cells[1, COL_HORARIO_ENTREGA].Value = "HORÁRIO DE ENTREGA";
            worksheet.Cells[1, COL_EMPRESA].Value = "POR QUAL EMPRESA";

            // Formatar cabeçalho
            using (var range = worksheet.Cells[1, 1, 1, COL_EMPRESA])
            {
                range.Style.Font.Bold = true;
                range.Style.Font.Color.SetColor(System.Drawing.Color.White);
                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 102, 204)); // Azul
                range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            }

            // Congelar a primeira linha
            worksheet.View.FreezePanes(2, 1);

            // Ajustar largura das colunas
            worksheet.Cells[1, 1, 1, COL_EMPRESA].AutoFitColumns();

            // Definir largura mínima para algumas colunas
            if (worksheet.Column(COL_PRODUTO).Width < 30)
                worksheet.Column(COL_PRODUTO).Width = 30;
            if (worksheet.Column(COL_EMPRESA).Width < 15)
                worksheet.Column(COL_EMPRESA).Width = 15;
        }

        private void CriarPlanilhaInicial()
        {
            try
            {
                lock (_lock)
                {
                    using (var package = new ExcelPackage())
                    {
                        var worksheet = package.Workbook.Worksheets.Add("Controle Pedidos");

                        // Criar cabeçalhos
                        CriarCabecalhos(worksheet);

                        // Adicionar linha de exemplo (opcional)
                        worksheet.Cells[2, COL_NUMERO_COTACAO].Value = "123456789";
                        worksheet.Cells[2, COL_PORTAL].Value = "Mercado Eletrônico";
                        worksheet.Cells[2, COL_CLIENTE].Value = "Nestle";
                        worksheet.Cells[2, COL_DATA_VENCIMENTO].Value = DateTime.Now.AddDays(7).ToString("dd/MM/yyyy");
                        worksheet.Cells[2, COL_HORARIO_VENCIMENTO].Value = "17:00";
                        worksheet.Cells[2, COL_PRODUTO].Value = "Exemplo de produto";
                        worksheet.Cells[2, COL_DATA_ENTREGA].Formula = $"=G{2}+7"; 
                        worksheet.Cells[2, COL_HORARIO_ENTREGA].Value = "";
                        worksheet.Cells[2, COL_EMPRESA].Value = "VENTURA";

                        // Formatar linha de exemplo em cinza claro
                        using (var range = worksheet.Cells[2, 1, 2, COL_EMPRESA])
                        {
                            range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        }

                        // Adicionar filtros
                        worksheet.Cells[1, 1, 1, COL_EMPRESA].AutoFilter = true;

                        // Salvar
                        var fileInfo = new FileInfo(_caminhoPlanilha);
                        package.SaveAs(fileInfo);

                        Console.WriteLine($"   ✅ Planilha de controle criada em: {_caminhoPlanilha}");
                        Console.WriteLine($"   📊 Colunas configuradas na ordem especificada");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"   ⚠️  Erro ao criar planilha inicial: {ex.Message}");
            }
        }

        public void AdicionarCotacaoNaPlanilha(Cotacao cotacao)
        {
            try
            {
                lock (_lock)
                {
                    using (var package = new ExcelPackage(new FileInfo(_caminhoPlanilha)))
                    {
                        var worksheet = package.Workbook.Worksheets[0];

                        // Encontrar a primeira linha vazia
                        int proximaLinha = 2;
                        while (proximaLinha <= 10000)
                        {
                            var valor = worksheet.Cells[proximaLinha, COL_NUMERO_COTACAO].Value?.ToString();
                            if (string.IsNullOrWhiteSpace(valor))
                            {
                                break;
                            }
                            proximaLinha++;
                        }

                        // Adicionar dados
                        worksheet.Cells[proximaLinha, COL_NUMERO_COTACAO].Value = cotacao.NumeroCotacao;
                        worksheet.Cells[proximaLinha, COL_PORTAL].Value = cotacao.Portal ?? "Mercado Eletrônico".ToUpper();
                        worksheet.Cells[proximaLinha, COL_CLIENTE].Value = cotacao.Empresa ?? "Nestle".ToUpper();

                        if (cotacao.DataVencimento != DateTime.MinValue)
                        {
                            worksheet.Cells[proximaLinha, COL_DATA_VENCIMENTO].Value = cotacao.DataVencimento;
                            worksheet.Cells[proximaLinha, COL_DATA_VENCIMENTO].Style.Numberformat.Format = "dd/MM/yyyy";
                        }

                        if (cotacao.HorarioVencimento != TimeSpan.Zero)
                        {
                            var horaDateTime = DateTime.Today.Add(cotacao.HorarioVencimento);
                            worksheet.Cells[proximaLinha, COL_HORARIO_VENCIMENTO].Value = horaDateTime;
                            worksheet.Cells[proximaLinha, COL_HORARIO_VENCIMENTO].Style.Numberformat.Format = "hh:mm";
                        }
                        string primeiroItem = cotacao.Itens != null && cotacao.Itens.Any()
                        ? (!string.IsNullOrWhiteSpace(cotacao.Itens[0]) ? cotacao.Itens[0] : "Produto não identificado")
                        : "Produto não identificado";
                        if (cotacao.Itens != null && cotacao.Itens.Any())
                        {
                            var produtosTexto = $"{primeiroItem} ({cotacao.Itens.Count} Itens)";
                            if (produtosTexto.Length > 32767)
                                produtosTexto = produtosTexto.Substring(0, 32760) + "...";
                            worksheet.Cells[proximaLinha, COL_PRODUTO].Value = produtosTexto;
                            worksheet.Cells[proximaLinha, COL_PRODUTO].Style.WrapText = true;
                        }

                        // Preservar fórmulas
                        if (string.IsNullOrEmpty(worksheet.Cells[proximaLinha, COL_DATA_ENTREGA].Formula))
                        {
                            worksheet.Cells[proximaLinha, COL_DATA_ENTREGA].Value = DateTime.Today;
                        }

                        if (string.IsNullOrEmpty(worksheet.Cells[proximaLinha, COL_HORARIO_ENTREGA].Formula))
                        {
                            worksheet.Cells[proximaLinha, COL_HORARIO_ENTREGA].Value = TimeOnly.FromDateTime(DateTime.Now);
                        }

                        worksheet.Cells[proximaLinha, COL_EMPRESA].Value = "VENTURA";

                        // Garantir AutoFilter
                        if (worksheet.Cells[1, 1, 1, COL_EMPRESA].AutoFilter == null)
                        {
                            worksheet.Cells[1, 1, 1, COL_EMPRESA].AutoFilter = true;
                        }

                        package.Save();
                    }
                }

                Console.WriteLine($"   ✅ Cotação {cotacao.NumeroCotacao} adicionada à planilha de controle");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"   ⚠️  Erro ao adicionar cotação na planilha: {ex.Message}");
                SalvarBackup(cotacao);
            }
        }

        private void SalvarBackup(Cotacao cotacao)
        {
            try
            {
                string backupPath = Path.Combine(Path.GetTempPath(), $"cotacao_backup_{DateTime.Now:yyyyMMdd_HHmmss}.txt");
                using (StreamWriter writer = new StreamWriter(backupPath))
                {
                    writer.WriteLine($"NÚMERO DA COTAÇÃO: {cotacao.NumeroCotacao}");
                    writer.WriteLine($"PORTAL: {cotacao.Portal}");
                    writer.WriteLine($"CLIENTE: {cotacao.Empresa}");
                    writer.WriteLine($"DATA DE VENCIMENTO: {cotacao.DataVencimento:dd/MM/yyyy}");
                    writer.WriteLine($"HORÁRIO DE VENCIMENTO: {cotacao.HorarioVencimento}");
                    writer.WriteLine($"PRODUTO: {(cotacao.Itens != null ? string.Join("; ", cotacao.Itens) : "")}");
                    writer.WriteLine($"DATA DE ENTREGA: ");
                    writer.WriteLine($"HORÁRIO DE ENTREGA: ");
                    writer.WriteLine($"POR QUAL EMPRESA: VENTURA");
                }
                Console.WriteLine($"   💾 Backup salvo em: {backupPath}");
            }
            catch { }
        }

        public void Dispose()
        {
            _fileWatcher?.Dispose();
        }
    }
}