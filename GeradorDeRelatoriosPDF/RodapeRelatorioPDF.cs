using iTextSharp.text.pdf;
using iTextSharp.text;

namespace GeradorDeRelatoriosPDF
{
    public class RodapeRelatorioPDF : PdfPageEventHelper
    {
        private readonly BaseFont familiaFonteRodape;
        private readonly Font fonteRodape;
        private readonly int totalPaginas;

        public RodapeRelatorioPDF(int totalPaginas)
        {
            familiaFonteRodape = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, false);
            fonteRodape = new Font(familiaFonteRodape, 8f, Font.NORMAL, BaseColor.Black);
            this.totalPaginas = totalPaginas;
        }

        public override void OnEndPage(PdfWriter writer, Document document)
        {
            base.OnEndPage(writer, document);
            AdicionarMomentoDaGeracao(writer, document);
            AdicionarNumeroPaginaAtual(writer, document);
        }

        private void AdicionarMomentoDaGeracao(PdfWriter writer, Document document)
        {
            string textoDataGeracao = $"Gerado em {DateTime.Now.ToShortDateString()} às {DateTime.Now.ToShortTimeString()}";

            PdfContentByte cb = writer.DirectContent;
            cb.BeginText();
            cb.SetFontAndSize(fonteRodape.BaseFont, fonteRodape.Size);
            cb.SetTextMatrix(document.LeftMargin, document.BottomMargin * 0.75f);
            cb.ShowText(textoDataGeracao);
            cb.EndText();
        }

        private void AdicionarNumeroPaginaAtual(PdfWriter writer, Document document)
        {
            int paginaAtual = writer.PageNumber;
            string textoPaginaAtual = "Página " + paginaAtual.ToString() + " de " + totalPaginas.ToString();

            PdfContentByte cb = writer.DirectContent;
            cb.BeginText();
            cb.SetFontAndSize(fonteRodape.BaseFont, fonteRodape.Size);
            cb.SetTextMatrix(document.PageSize.Width - (document.RightMargin + fonteRodape.BaseFont.GetWidthPoint(textoPaginaAtual, fonteRodape.Size)), document.BottomMargin * 0.75f);
            cb.ShowText(textoPaginaAtual);
            cb.EndText();
        }
    }
}