
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Shapes;
using MigraDoc.DocumentObjectModel.Tables;

namespace PDFLab
{
    public class BizMigraDocSample
    {

    }

    public class MigraDocSample_HelloWorld
    {
        /// <summary>
        /// Creates an absolutely minimalistic document.
        /// </summary>
        public static Document CreateDocument()
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
    }

    /// <summary>
    /// This sample shows how to create a simple invoice of a fictional book store. The invoice document is created with the MigraDoc document object model and then rendered to PDF with PDFsharp.
    /// </summary>
    public class MigraDocSample_CreateInvoiceDocument
    {
        // attributes
        protected Document _document = null;

        // document content
        protected TextFrame _addressFrame = null;
        protected Table _table = null;

        // skin attributes
        protected readonly Color TABLE_BORDER_COLOR = Color.FromCmyk(50, 100, 0, 0);
        protected readonly Color TABLE_BLUE_COLOR = Color.FromCmyk(29, 15, 0, 0);
        protected readonly Color TABLE_GRAY_COLOR = Color.FromCmyk(9, 5, 0, 0);

        // data
        protected struct InvoiceItemInfo
        {
            public int itemNumber;
            public string title;
            public string author;
            public double quantity;
            public double price;
            public double discount;

            public InvoiceItemInfo(int itemNumber,
                                string title,
                                string author,
                                double quantity,
                                double price,
                                double discount)
            {
                this.itemNumber = itemNumber;
                this.title = title;
                this.author = author;
                this.quantity = quantity;
                this.price = price;
                this.discount = discount;
            }
        }

        protected InvoiceItemInfo [] _invoiceList = new InvoiceItemInfo[] {
                new InvoiceItemInfo(1, "PDF Reference Version 1.6 (5th Edition)", "Adobe Systems Inc.", 2, 37, 5),
                new InvoiceItemInfo(1, "iText in Action - Creating and Manipulating PDF", "Bruno Lowagie", 1, 36.95, 5 ),
                new InvoiceItemInfo(2, "The C# Programming Language", "Anders Hejlsberg, Peter Golde, Scott",5, 23, 10 ),
                new InvoiceItemInfo(3, "The Elegant Universe: Superstrings, Hidden Dimensions, and the Quest for the Ultimate Theory", "Brian Greene", 1, 17.9, 0),
                new InvoiceItemInfo(4, "'Surely You're Joking, Mr. Feynman!' (Adventures of a Curious Character)", "Richard P. Feynman, Ralph Leighton", 3, 9.8, 0),
                new InvoiceItemInfo(5, "The Mind's I", "Douglas R. Hofstadter, Daniel C. Dennett", 1, 29.9, 0),
                new InvoiceItemInfo(6, "The Blind Watchmaker: Why the Evidence of Evolution Reveals a Universe Without Design", "Richard Dawkins", 1, 11, 0),
                new InvoiceItemInfo(7, "How the Mind Works", "Steven Pinker", 1, 16.3, 0),
                new InvoiceItemInfo(8, "The Origins of Life: From the Birth of Life to the Origin of Language", "John Maynard Smith, Eörs Szathmáry", 1, 29.45, 0),
            };

        public Document CreateDocument()
        {
            // Create a new MigraDoc document
            _document = new Document();
            _document.Info.Title = "A sample invoice";
            _document.Info.Subject = "Demonstrates how to create an invoice.";
            _document.Info.Author = "Stefan Lange";

            DefineStyles();

            CreatePage();

            FillContent();

            return _document;
        }

        /// <summary>
        /// Styles define how the text will look:
        /// </summary>
        protected void DefineStyles()
        {
            // Get the predefined style Normal.
            Style style = _document.Styles["Normal"];
            // Because all styles are derived from Normal, the next line changes the
            // font of the whole document. Or, more exactly, it changes the font of
            // all styles and paragraphs that do not redefine the font.
            style.Font.Name = "Verdana";

            style = _document.Styles[StyleNames.Header];
            style.ParagraphFormat.AddTabStop("16cm", TabAlignment.Right);

            style = _document.Styles[StyleNames.Footer];
            style.ParagraphFormat.AddTabStop("8cm", TabAlignment.Center);

            // Create a new style called Table based on style Normal
            style = _document.Styles.AddStyle("Table", "Normal");
            style.Font.Name = "Verdana";
            style.Font.Name = "Times New Roman";
            style.Font.Size = 9;

            // Create a new style called Reference based on style Normal
            style = _document.Styles.AddStyle("Reference", "Normal");
            style.ParagraphFormat.SpaceBefore = "5mm";
            style.ParagraphFormat.SpaceAfter = "5mm";
            style.ParagraphFormat.TabStops.AddTabStop("16cm", TabAlignment.Right);
        }

        /// <summary>
        /// Create the page with invoice table, header, footer:
        /// </summary>
        protected void CreatePage()
        {
            // Each MigraDoc document needs at least one section.
            Section section = _document.AddSection();

            // Put a logo in the header
            Image image = section.Headers.Primary.AddImage("imgs\\CustomerIcon3.jpg");
            image.Height = "2.5cm";
            image.LockAspectRatio = true;
            image.RelativeVertical = RelativeVertical.Line;
            image.RelativeHorizontal = RelativeHorizontal.Margin;
            image.Top = ShapePosition.Top;
            image.Left = ShapePosition.Right;
            image.WrapFormat.Style = WrapStyle.Through;

            // Create footer
            Paragraph paragraph = section.Footers.Primary.AddParagraph();
            paragraph.AddText("PowerBooks Inc · Sample Street 42 · 56789 Cologne · Germany");
            paragraph.Format.Font.Size = 9;
            paragraph.Format.Alignment = ParagraphAlignment.Center;

            // Create the text frame for the address
            _addressFrame = section.AddTextFrame();
            _addressFrame.Height = "3.0cm";
            _addressFrame.Width = "7.0cm";
            _addressFrame.Left = ShapePosition.Left;
            _addressFrame.RelativeHorizontal = RelativeHorizontal.Margin;
            _addressFrame.Top = "5.0cm";
            _addressFrame.RelativeVertical = RelativeVertical.Page;

            // Put sender in address frame
            paragraph = _addressFrame.AddParagraph("PowerBooks Inc · Sample Street 42 · 56789 Cologne");
            paragraph.Format.Font.Name = "Times New Roman";
            paragraph.Format.Font.Size = 7;
            paragraph.Format.SpaceAfter = 3;

            // Add the print date field
            paragraph = section.AddParagraph();
            paragraph.Format.SpaceBefore = "8cm";
            paragraph.Style = "Reference";
            paragraph.AddFormattedText("INVOICE", TextFormat.Bold);
            paragraph.AddTab();
            paragraph.AddText("Cologne, ");
            paragraph.AddDateField("dd.MM.yyyy");

            // Create the item table
            _table = section.AddTable();
            _table.Style = "Table";
            _table.Borders.Color = TABLE_BORDER_COLOR;
            _table.Borders.Width = 0.25;
            _table.Borders.Left.Width = 0.5;
            _table.Borders.Right.Width = 0.5;
            _table.Rows.LeftIndent = 0;

            // Before you can add a row, you must define the columns
            Column column = _table.AddColumn("1cm");
            column.Format.Alignment = ParagraphAlignment.Center;

            column = _table.AddColumn("2.5cm");
            column.Format.Alignment = ParagraphAlignment.Right;

            column = _table.AddColumn("3cm");
            column.Format.Alignment = ParagraphAlignment.Right;

            column = _table.AddColumn("3.5cm");
            column.Format.Alignment = ParagraphAlignment.Right;

            column = _table.AddColumn("2cm");
            column.Format.Alignment = ParagraphAlignment.Center;

            column = _table.AddColumn("4cm");
            column.Format.Alignment = ParagraphAlignment.Right;

            // Create the header of the table
            Row row = _table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = TABLE_BLUE_COLOR;
            row.Cells[0].AddParagraph("Item");
            row.Cells[0].Format.Font.Bold = false;
            row.Cells[0].Format.Alignment = ParagraphAlignment.Left;
            row.Cells[0].VerticalAlignment = VerticalAlignment.Bottom;
            row.Cells[0].MergeDown = 1;
            row.Cells[1].AddParagraph("Title and Author");
            row.Cells[1].Format.Alignment = ParagraphAlignment.Left;
            row.Cells[1].MergeRight = 3;
            row.Cells[5].AddParagraph("Extended Price");
            row.Cells[5].Format.Alignment = ParagraphAlignment.Left;
            row.Cells[5].VerticalAlignment = VerticalAlignment.Bottom;
            row.Cells[5].MergeDown = 1;

            row = _table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = TABLE_BLUE_COLOR;
            row.Cells[1].AddParagraph("Quantity");
            row.Cells[1].Format.Alignment = ParagraphAlignment.Left;
            row.Cells[2].AddParagraph("Unit Price");
            row.Cells[2].Format.Alignment = ParagraphAlignment.Left;
            row.Cells[3].AddParagraph("Discount (%)");
            row.Cells[3].Format.Alignment = ParagraphAlignment.Left;
            row.Cells[4].AddParagraph("Taxable");
            row.Cells[4].Format.Alignment = ParagraphAlignment.Left;

            _table.SetEdge(0, 0, 6, 2, Edge.Box, BorderStyle.Single, 0.75, Color.Empty);
        }

        /// <summary>
        /// This routine adds the dynamic data to the invoice:
        /// </summary>
        protected void FillContent()
        {
            // Fill address in address text frame
            // XPathNavigator item = SelectItem("/invoice/to");
            Paragraph paragraph = _addressFrame.AddParagraph();
            paragraph.AddText(@"name/singleName");
            paragraph.AddLineBreak();
            paragraph.AddText(@"address/line1");
            paragraph.AddLineBreak();
            paragraph.AddText("address/postalCode address/city");

            // Iterate the invoice items
            double totalExtendedPrice = 0;
            //XPathNodeIterator iter = this.navigator.Select("/invoice/items/*");
            var iter = _invoiceList.GetEnumerator();
            while (iter.MoveNext())
            {
                InvoiceItemInfo item = (InvoiceItemInfo)iter.Current;

                // Each item fills two rows
                Row row1 = _table.AddRow();
                Row row2 = _table.AddRow();
                row1.TopPadding = 1.5;
                row1.Cells[0].Shading.Color = TABLE_GRAY_COLOR;
                row1.Cells[0].VerticalAlignment = VerticalAlignment.Center;
                row1.Cells[0].MergeDown = 1;
                row1.Cells[1].Format.Alignment = ParagraphAlignment.Left;
                row1.Cells[1].MergeRight = 3;
                row1.Cells[5].Shading.Color = TABLE_GRAY_COLOR;
                row1.Cells[5].MergeDown = 1;

                row1.Cells[0].AddParagraph(item.itemNumber.ToString()); // itemNumber

                paragraph = row1.Cells[1].AddParagraph();
                paragraph.AddFormattedText(item.title); // title
                paragraph.AddFormattedText(" by ", TextFormat.Italic);
                paragraph.AddText(item.author); // author

                row2.Cells[1].AddParagraph(item.quantity.ToString()); // quantity

                row2.Cells[2].AddParagraph(item.price.ToString("0.00") + " €"); // price

                row2.Cells[3].AddParagraph(item.discount.ToString("0.0")); // discount

                row2.Cells[4].AddParagraph();

                //row2.Cells[5].AddParagraph(item.price.ToString("0.00")); // price

                double extendedPrice = item.quantity * item.price;
                extendedPrice = extendedPrice * (100 - item.discount) / 100;
                row1.Cells[5].AddParagraph(extendedPrice.ToString("0.00") + " €"); // Extended Price
                row1.Cells[5].VerticalAlignment = VerticalAlignment.Bottom;

                totalExtendedPrice += extendedPrice;

                _table.SetEdge(0, _table.Rows.Count - 2, 6, 2, Edge.Box, BorderStyle.Single, 0.75);
            }

            // Add an invisible row as a space line to the table
            Row row = _table.AddRow();
            row.Borders.Visible = false;

            // Add the total price row
            row = _table.AddRow();
            row.Cells[0].Borders.Visible = false;
            row.Cells[0].AddParagraph("Total Price");
            row.Cells[0].Format.Font.Bold = true;
            row.Cells[0].Format.Alignment = ParagraphAlignment.Right;
            row.Cells[0].MergeRight = 4;
            row.Cells[5].AddParagraph(totalExtendedPrice.ToString("0.00") + " €");

            // Add the VAT row
            row = _table.AddRow();
            row.Cells[0].Borders.Visible = false;
            row.Cells[0].AddParagraph("VAT (19%)");
            row.Cells[0].Format.Font.Bold = true;
            row.Cells[0].Format.Alignment = ParagraphAlignment.Right;
            row.Cells[0].MergeRight = 4;
            row.Cells[5].AddParagraph((0.19 * totalExtendedPrice).ToString("0.00") + " €");

            // Add the additional fee row
            row = _table.AddRow();
            row.Cells[0].Borders.Visible = false;
            row.Cells[0].AddParagraph("Shipping and Handling");
            row.Cells[5].AddParagraph(0.ToString("0.00") + " €");
            row.Cells[0].Format.Font.Bold = true;
            row.Cells[0].Format.Alignment = ParagraphAlignment.Right;
            row.Cells[0].MergeRight = 4;

            // Add the total due row
            row = _table.AddRow();
            row.Cells[0].AddParagraph("Total Due");
            row.Cells[0].Borders.Visible = false;
            row.Cells[0].Format.Font.Bold = true;
            row.Cells[0].Format.Alignment = ParagraphAlignment.Right;
            row.Cells[0].MergeRight = 4;
            totalExtendedPrice += 0.19 * totalExtendedPrice;
            row.Cells[5].AddParagraph(totalExtendedPrice.ToString("0.00") + " €");

            // Set the borders of the specified cell range
            _table.SetEdge(5, _table.Rows.Count - 4, 1, 4, Edge.Box, BorderStyle.Single, 0.75);

            // Add the notes paragraph
            paragraph = _document.LastSection.AddParagraph();
            paragraph.Format.SpaceBefore = "1cm";
            paragraph.Format.Borders.Width = 0.75;
            paragraph.Format.Borders.Distance = 3;
            paragraph.Format.Borders.Color = TABLE_BLUE_COLOR;
            paragraph.Format.Shading.Color = TABLE_GRAY_COLOR;
            paragraph.AddText(@"This is a sample invoice created with MigraDoc. The PDF document is rendered with PDFsharp.");
        }

        /// <summary>
        /// This routine adds the dynamic data to the invoice:
        /// </summary>
        //protected static void FillContent(ref Document document)
        //{
        //    // Fill address in address text frame
        //    XPathNavigator item = SelectItem("/invoice/to");
        //    Paragraph paragraph = this.addressFrame.AddParagraph();
        //    paragraph.AddText(GetValue(item, "name/singleName"));
        //    paragraph.AddLineBreak();
        //    paragraph.AddText(GetValue(item, "address/line1"));
        //    paragraph.AddLineBreak();
        //    paragraph.AddText(GetValue(item, "address/postalCode") + " " + GetValue(item, "address/city"));
        //
        //    // Iterate the invoice items
        //    double totalExtendedPrice = 0;
        //    XPathNodeIterator iter = this.navigator.Select("/invoice/items/*");
        //
        //    while (iter.MoveNext())
        //    {
        //        item = iter.Current;
        //        double quantity = GetValueAsDouble(item, "quantity");
        //        double price = GetValueAsDouble(item, "price");
        //        double discount = GetValueAsDouble(item, "discount");
        //
        //        // Each item fills two rows
        //        Row row1 = this.table.AddRow();
        //        Row row2 = this.table.AddRow();
        //        row1.TopPadding = 1.5;
        //        row1.Cells[0].Shading.Color = TableGray;
        //        row1.Cells[0].VerticalAlignment = VerticalAlignment.Center;
        //        row1.Cells[0].MergeDown = 1;
        //        row1.Cells[1].Format.Alignment = ParagraphAlignment.Left;
        //        row1.Cells[1].MergeRight = 3;
        //        row1.Cells[5].Shading.Color = TableGray;
        //        row1.Cells[5].MergeDown = 1;
        //
        //        row1.Cells[0].AddParagraph(GetValue(item, "itemNumber"));
        //        paragraph = row1.Cells[1].AddParagraph();
        //        paragraph.AddFormattedText(GetValue(item, "title"), TextFormat.Bold);
        //        paragraph.AddFormattedText(" by ", TextFormat.Italic);
        //        paragraph.AddText(GetValue(item, "author"));
        //        row2.Cells[1].AddParagraph(GetValue(item, "quantity"));
        //        row2.Cells[2].AddParagraph(price.ToString("0.00") + " €");
        //        row2.Cells[3].AddParagraph(discount.ToString("0.0"));
        //        row2.Cells[4].AddParagraph();
        //        row2.Cells[5].AddParagraph(price.ToString("0.00"));
        //        double extendedPrice = quantity * price;
        //        extendedPrice = extendedPrice* (100 - discount) / 100;
        //        row1.Cells[5].AddParagraph(extendedPrice.ToString("0.00") + " €");
        //        row1.Cells[5].VerticalAlignment = VerticalAlignment.Bottom;
        //        totalExtendedPrice += extendedPrice;
        //
        //        this.table.SetEdge(0, this.table.Rows.Count - 2, 6, 2, Edge.Box, BorderStyle.Single, 0.75);
        //    }
        //
        //    // Add an invisible row as a space line to the table
        //    Row row = this.table.AddRow();
        //    row.Borders.Visible = false;
        //
        //    // Add the total price row
        //    row = this.table.AddRow();
        //    row.Cells [0].Borders.Visible = false;
        //    row.Cells [0].AddParagraph("Total Price");
        //    row.Cells [0].Format.Font.Bold = true;
        //    row.Cells [0].Format.Alignment = ParagraphAlignment.Right;
        //    row.Cells [0].MergeRight = 4;
        //    row.Cells [5].AddParagraph(totalExtendedPrice.ToString("0.00") + " €");
        //
        //    // Add the VAT row
        //    row = this.table.AddRow();
        //    row.Cells [0].Borders.Visible = false;
        //    row.Cells [0].AddParagraph("VAT (19%)");
        //    row.Cells [0].Format.Font.Bold = true;
        //    row.Cells [0].Format.Alignment = ParagraphAlignment.Right;
        //    row.Cells [0].MergeRight = 4;
        //    row.Cells [5].AddParagraph((0.19 * totalExtendedPrice).ToString("0.00") + " €");
        //
        //    // Add the additional fee row
        //    row = this.table.AddRow();
        //    row.Cells [0].Borders.Visible = false;
        //    row.Cells [0].AddParagraph("Shipping and Handling");
        //    row.Cells [5].AddParagraph(0.ToString("0.00") + " €");
        //    row.Cells [0].Format.Font.Bold = true;
        //    row.Cells [0].Format.Alignment = ParagraphAlignment.Right;
        //    row.Cells [0].MergeRight = 4;
        //
        //    // Add the total due row
        //    row = this.table.AddRow();
        //    row.Cells [0].AddParagraph("Total Due");
        //    row.Cells [0].Borders.Visible = false;
        //    row.Cells [0].Format.Font.Bold = true;
        //    row.Cells [0].Format.Alignment = ParagraphAlignment.Right;
        //    row.Cells [0].MergeRight = 4;
        //    totalExtendedPrice += 0.19 * totalExtendedPrice;
        //    row.Cells [5].AddParagraph(totalExtendedPrice.ToString("0.00") + " €");
        //
        //    // Set the borders of the specified cell range
        //    this.table.SetEdge(5, this.table.Rows.Count - 4, 1, 4, Edge.Box, BorderStyle.Single, 0.75);
        //
        //    // Add the notes paragraph
        //    paragraph = this.document.LastSection.AddParagraph();
        //    paragraph.Format.SpaceBefore = "1cm";
        //    paragraph.Format.Borders.Width = 0.75;
        //    paragraph.Format.Borders.Distance = 3;
        //    paragraph.Format.Borders.Color = TableBorder;
        //    paragraph.Format.Shading.Color = TableGray;
        //    item = SelectItem("/invoice");
        //    paragraph.AddText(GetValue(item, "notes"));
        //}
    }
}
