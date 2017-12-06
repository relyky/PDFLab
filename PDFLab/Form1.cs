using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;
using MigraDoc.DocumentObjectModel.Shapes;
using MigraDoc.DocumentObjectModel.Tables;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
//using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PDFLab
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new PDF document
            PdfDocument document = new PdfDocument();
            document.Info.Title = "Created with PDFsharp";

            // Create an empty page
            PdfPage page = document.AddPage();

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            // Create a font
            XFont font = new XFont("Verdana", 20, XFontStyle.BoldItalic);

            // Draw the text
            gfx.DrawString("Hello, World!", font, XBrushes.Black,
            new XRect(0, 0, page.Width, page.Height),
            XStringFormats.Center);

            // Draw the line, it could used to draw table
            XPen pen = new XPen(XColor.FromKnownColor(XKnownColor.DeepPink), 1.0);
            gfx.DrawLine(pen, 45, 250, 45, 703);
            gfx.DrawLine(pen, 87, 250, 87, 703);
            gfx.DrawLine(pen, 150, 250, 150, 703);
            gfx.DrawLine(pen, 291, 250, 291, 703);
            gfx.DrawLine(pen, 381, 250, 381, 703);
            gfx.DrawLine(pen, 461, 250, 461, 703);
            gfx.DrawLine(pen, 571, 250, 571, 703);

            // Save the document...
            const string filename = "d:\\HelloWorld.pdf";
            document.Save(filename);

            // ...and start a viewer.
            System.Diagnostics.Process.Start(filename);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // Create a MigraDoc document
            Document document = MigraDocSample_HelloWorld.CreateDocument();
            document.UseCmykColor = true;
 
            // ===== Unicode encoding and font program embedding in MigraDoc is demonstrated here =====
            // A flag indicating whether to create a Unicode PDF or a WinAnsi PDF file.
            // This setting applies to all fonts used in the PDF document.
            // This setting has no effect on the RTF renderer.
            const bool unicode = false;

            // An enum indicating whether to embed fonts or not.
            // This setting applies to all font programs used in the document.
            // This setting has no effect on the RTF renderer.
            // (The term 'font program' is used by Adobe for a file containing a font. Technically a 'font file'
            // is a collection of small programs and each program renders the glyph of a character when executed.
            // Using a font in PDFsharp may lead to the embedding of one or more font programms, because each outline
            // (regular, bold, italic, bold+italic, ...) has its own fontprogram)
            const PdfFontEmbedding embedding = PdfFontEmbedding.Always;
 
            // ========================================================================================
            // Create a renderer for the MigraDoc document.
            PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(unicode, embedding);

            // Associate the MigraDoc document with a renderer
            pdfRenderer.Document = document;

            // Layout and render document to PDF
            pdfRenderer.RenderDocument();

            // Save the document...
            const string filename = "d:\\HelloWorld.pdf";
            pdfRenderer.PdfDocument.Save(filename);

            // ...and start a viewer.
            System.Diagnostics.Process.Start(filename);
        }

        #region BizMigraDoc_CreateHelloDocument

        /// <summary>
        /// Creates an absolutely minimalistic document.
        /// </summary>
        private static Document BizMigraDoc_CreateHelloDocument()
        {
            // Create a new MigraDoc document
            Document document = new Document();
 
            // Add a section to the document
            Section section = document.AddSection();
 
            // Add a paragraph to the section
            Paragraph paragraph = section.AddParagraph();
 
            paragraph.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 30, 20, 50);
 
            // Add some text to the paragraph
            paragraph.AddFormattedText("Hello, World!", TextFormat.Bold);
 
            return document;
        }

        #endregion

        private void button3_Click(object sender, EventArgs e)
        {
            // Create a MigraDoc document
            MigraDocSample_CreateInvoiceDocument invoiceSample = new MigraDocSample_CreateInvoiceDocument();
            Document document = invoiceSample.CreateDocument();
            document.UseCmykColor = true;

            // Create a renderer for the MigraDoc document.
            const bool unicode = false;
            const PdfFontEmbedding embedding = PdfFontEmbedding.Always;
            PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(unicode, embedding);

            // Associate the MigraDoc document with a renderer
            pdfRenderer.Document = document;

            // Layout and render document to PDF
            pdfRenderer.RenderDocument();

            // Save the document...
            const string filename = "d:\\HelloWorld.pdf";
            pdfRenderer.PdfDocument.Save(filename);

            // ...and start a viewer.
            System.Diagnostics.Process.Start(filename);
        }


}
}
