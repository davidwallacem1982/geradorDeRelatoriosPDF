using GeradorDeRelatoriosPDF.Models;
using iTextSharp.text.pdf;
using iTextSharp.text;
using System.Diagnostics;
using System.Text.Json;
using GeradorDeRelatoriosPDF;

internal class Program
{
    static List<Pessoa> pessoas = new List<Pessoa>();
    private static void Main(string[] args)
    {
        DesserializarPessoas();
        GerarRelatorioPDF(100);
    }
    public static void DesserializarPessoas()
    {
        try
        {
            var filePath = "pessoas.json";
            if (File.Exists(filePath))
            {
                var dados = File.ReadAllText(filePath);
                pessoas = JsonSerializer.Deserialize<List<Pessoa>>(dados);
            }
        }
        catch (Exception ex)
        {
            // Lidar com a exceção de alguma forma (ex: registro de log, relançamento, etc.)
            Console.WriteLine("Erro ao desserializar pessoas: " + ex.Message);
        }
    }

    public static void GerarRelatorioPDF(int qtdePessoas)
    {
        var pessoasSelecionadas = pessoas.Take(qtdePessoas).ToList();
        if (pessoasSelecionadas.Count > 0)
        {
            var qtdeRegistrosPrimeiraPagina = 24;
            var qtdeRegistrosPorPagina = 29;

            var totalRegistros = pessoasSelecionadas.Count;
            var qtdeRegistrosRestantes = totalRegistros - qtdeRegistrosPrimeiraPagina;
            var totalPaginas = 1 + (int)Math.Ceiling((float)qtdeRegistrosRestantes / qtdeRegistrosPorPagina);

            var pxPorMm = 72 / 25.2F;
            var pdf = new Document(PageSize.A4, 15 * pxPorMm, 15 * pxPorMm, 15 * pxPorMm, 20 * pxPorMm);

            var nomeArquivo = $"pessoas.{DateTime.Now.ToString("yyyyMMddHHmmss")}.pdf";

            using (var arquivo = new FileStream(nomeArquivo, FileMode.Create))
            {
                using (var writer = PdfWriter.GetInstance(pdf, arquivo))
                {
                    writer.PageEvent = new RodapeRelatorioPDF(totalPaginas);
                    pdf.Open();

                    var fonteBase = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, false);
                    var fonteParagrafo = new Font(fonteBase, 32, Font.NORMAL, BaseColor.Black);
                    var fonteLink = new Font(fonteBase, 9.9f, Font.NORMAL, BaseColor.Blue);

                    var titulo = new Paragraph("Relatório de Funcionários\n\n", fonteParagrafo);
                    titulo.Alignment = Element.ALIGN_LEFT;
                    titulo.SpacingAfter = 4;
                    pdf.Add(titulo);

                    var link = new Chunk("Linkedin", fonteLink);
                    var larguraTexto = fonteBase.GetWidthPoint(link.Content, fonteLink.Size);
                    link.SetAnchor("https://www.linkedin.com/in/david-wallace-marques-ferreira/");
                    var caixaTexto = new ColumnText(writer.DirectContent);
                    caixaTexto.AddElement(link);
                    caixaTexto.SetSimpleColumn(
                        pdf.PageSize.Width - pdf.RightMargin - larguraTexto,
                        pdf.PageSize.Height - pdf.TopMargin - (30 * pxPorMm),
                        pdf.PageSize.Width - pdf.RightMargin,
                        pdf.PageSize.Height - pdf.TopMargin - (18 * pxPorMm));
                    caixaTexto.Go();

                    var caminhoImagem = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "img", "thankYou.png");
                    if (File.Exists(caminhoImagem))
                    {
                        Image logo = Image.GetInstance(caminhoImagem);
                        float razaoLarguraAltura = logo.Width / logo.Height;
                        float alturaLogo = 70;
                        float larguraLogo = alturaLogo * razaoLarguraAltura;
                        logo.ScaleToFit(larguraLogo, alturaLogo);
                        var margemEsquerda = pdf.PageSize.Width - pdf.RightMargin - larguraLogo;
                        var margemTopo = pdf.PageSize.Height - pdf.TopMargin - 54;
                        logo.SetAbsolutePosition(margemEsquerda, margemTopo);
                        writer.DirectContent.AddImage(logo, false);
                    }

                    var tabela = new PdfPTable(5);
                    float[] larguras = { 0.6f, 2f, 1.5f, 1f, 1f };
                    tabela.SetWidths(larguras);
                    tabela.DefaultCell.BorderWidth = 0;
                    tabela.WidthPercentage = 100;

                    CriarCelulaTexto(tabela, "Código", PdfPCell.ALIGN_CENTER, true);
                    CriarCelulaTexto(tabela, "Nome", PdfPCell.ALIGN_LEFT, true);
                    CriarCelulaTexto(tabela, "Profissão", PdfPCell.ALIGN_CENTER, true);
                    CriarCelulaTexto(tabela, "Salário", PdfPCell.ALIGN_CENTER, true);
                    CriarCelulaTexto(tabela, "Empregada", PdfPCell.ALIGN_CENTER, true);

                    foreach (var pessoa in pessoasSelecionadas)
                    {
                        CriarCelulaTexto(tabela, pessoa.IdPessoa.ToString("D6"), PdfPCell.ALIGN_CENTER);
                        CriarCelulaTexto(tabela, pessoa.Nome + " " + pessoa.Sobrenome);
                        CriarCelulaTexto(tabela, pessoa.Profissao.Nome, PdfPCell.ALIGN_CENTER);
                        CriarCelulaTexto(tabela, pessoa.Salario.ToString("C2"), PdfPCell.ALIGN_RIGHT);
                        var caminhoImagemCelula = pessoa.Empregado ? "img/ok.png" : "img/nok.png";
                        caminhoImagemCelula = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, caminhoImagemCelula);
                        CriarCelulaImagem(tabela, caminhoImagemCelula, 20, 20);
                    }

                    pdf.Add(tabela);
                    pdf.Close();

                    var caminhoPDF = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, nomeArquivo);
                    if (File.Exists(caminhoPDF))
                    {
                        Process.Start(new ProcessStartInfo()
                        {
                            FileName = caminhoPDF,
                            UseShellExecute = true
                        });
                    }
                }
            }
        }
        else
        {
            Console.WriteLine("Nenhum produto foi retornado.");
        }
    }

    public static void CriarCelulaTexto(PdfPTable tabela, string texto,
    int alinhamento = PdfPCell.ALIGN_LEFT,
    bool negrito = false, bool italico = false,
    int tamanhoFonte = 12, int alturaCelula = 25)
    {
        var estilo = GetFontStyle(negrito, italico);
        var fonte = GetFont(estilo, tamanhoFonte);

        var bgColor = GetBackgroundColor(tabela.Rows.Count);

        var celula = new PdfPCell(new Phrase(texto, fonte))
        {
            HorizontalAlignment = alinhamento,
            VerticalAlignment = PdfPCell.ALIGN_MIDDLE,
            Border = 0,
            BorderWidthBottom = 1,
            PaddingBottom = 5,
            FixedHeight = alturaCelula,
            BackgroundColor = bgColor
        };

        tabela.AddCell(celula);
    }

    private static int GetFontStyle(bool negrito, bool italico)
    {
        if (negrito && italico)
        {
            return Font.BOLDITALIC;
        }
        else if (negrito)
        {
            return Font.BOLD;
        }
        else if (italico)
        {
            return Font.ITALIC;
        }
        else
        {
            return Font.NORMAL;
        }
    }

    public static void CriarCelulaImagem(PdfPTable tabela, string caminhoImagem,
    int larguraImagem, int alturaImagem, int alturaCelula = 25)
    {
        var bgColor = GetBackgroundColor(tabela.Rows.Count);

        if (File.Exists(caminhoImagem))
        {
            var imagem = Image.GetInstance(caminhoImagem);
            imagem.ScaleToFit(larguraImagem, alturaImagem);

            var celula = new PdfPCell(imagem)
            {
                HorizontalAlignment = PdfPCell.ALIGN_CENTER,
                VerticalAlignment = PdfPCell.ALIGN_MIDDLE,
                Border = 0,
                BorderWidthBottom = 1,
                FixedHeight = alturaCelula,
                BackgroundColor = bgColor
            };

            tabela.AddCell(celula);
        }
        else
        {
            CriarCelulaTexto(tabela, "ERRO", PdfPCell.ALIGN_CENTER, false, false, 12, alturaCelula);
        }
    }

    private static Font GetFont(int estilo, int tamanhoFonte)
    {
        var fonteBase = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, false);
        return new Font(fonteBase, tamanhoFonte, estilo, BaseColor.Black);
    }

    private static BaseColor GetBackgroundColor(int rowCount)
    {
        return rowCount % 2 == 1 ? new BaseColor(0.95f, 0.95f, 0.95f) : BaseColor.White;
    }
}