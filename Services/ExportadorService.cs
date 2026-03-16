using MercadoEletronico.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace MercadoEletronico.Services
{
    public class ExportadorService
    {
        public void ExportarParaExcel(List<Cotacao> cotacoes)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using var package = new ExcelPackage();
                var worksheet = package.Workbook.Worksheets.Add("Cotações");

                // Cabeçalhos
                string[] headers = {
                    "Número Cotação", "Portal", "Empresa", "Data Vencimento",
                    "Horário Vencimento", "Data Registro", "Horário Registro",
                    "Solicitante", "Status", "Vendedor"
                };

                for (int i = 0; i < headers.Length; i++)
                {
                    var cell = worksheet.Cells[1, i + 1];
                    cell.Value = headers[i];
                    cell.Style.Font.Bold = true;
                    cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                }

                // Dados
                for (int i = 0; i < cotacoes.Count; i++)
                {
                    var cotacao = cotacoes[i];
                    var row = i + 2;

                    worksheet.Cells[row, 1].Value = cotacao.NumeroCotacao;
                    worksheet.Cells[row, 2].Value = cotacao.Portal;
                    worksheet.Cells[row, 3].Value = cotacao.Empresa;
                    worksheet.Cells[row, 4].Value = cotacao.DataVencimento;
                    worksheet.Cells[row, 4].Style.Numberformat.Format = "dd/MM/yyyy";
                    worksheet.Cells[row, 5].Value = cotacao.HorarioVencimento;
                    worksheet.Cells[row, 6].Value = cotacao.DataRegistro;
                    worksheet.Cells[row, 6].Style.Numberformat.Format = "dd/MM/yyyy";
                    worksheet.Cells[row, 7].Value = cotacao.HorarioRegistro;
                    worksheet.Cells[row, 8].Value = cotacao.Solicitante;
                    worksheet.Cells[row, 9].Value = cotacao.Status;
                    worksheet.Cells[row, 10].Value = cotacao.Vendedor;
                }

                // Formatação
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                var borderStyle = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                worksheet.Cells[worksheet.Dimension.Address].Style.Border.Top.Style = borderStyle;
                worksheet.Cells[worksheet.Dimension.Address].Style.Border.Bottom.Style = borderStyle;
                worksheet.Cells[worksheet.Dimension.Address].Style.Border.Left.Style = borderStyle;
                worksheet.Cells[worksheet.Dimension.Address].Style.Border.Right.Style = borderStyle;

                // Salvar
                var fileName = $"Cotações_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
                var filePath = Path.Combine(Directory.GetCurrentDirectory(), fileName);

                package.SaveAs(new FileInfo(filePath));

                Console.WriteLine($"\n📁 Arquivo Excel salvo: {filePath}");
                Console.WriteLine($"📊 Total de registros: {cotacoes.Count}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\n⚠️  Erro ao exportar Excel: {ex.Message}");
                ExportarParaCSV(cotacoes);
            }
        }

        private void ExportarParaCSV(List<Cotacao> cotacoes)
        {
            try
            {
                var lines = new List<string>
                {
                    "Número Cotação;Portal;Empresa;Data Vencimento;Horário Vencimento;Data Registro;Horário Registro;Solicitante;Status;Vendedor"
                };

                lines.AddRange(cotacoes.Select(c =>
                    $"{c.NumeroCotacao};{c.Portal};{c.Empresa};{c.DataVencimento:dd/MM/yyyy};{c.HorarioVencimento};" +
                    $"{c.DataRegistro:dd/MM/yyyy};{c.HorarioRegistro};{c.Solicitante};{c.Status};{c.Vendedor}"));

                var fileName = $"Cotações_{DateTime.Now:yyyyMMdd_HHmmss}.csv";
                File.WriteAllLines(fileName, lines);
                Console.WriteLine($"\n📁 Arquivo CSV salvo: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\n❌ Erro ao exportar CSV: {ex.Message}");
            }
        }
    }
}