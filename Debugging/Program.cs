using System;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;
using Elistia.DotNetRtfWriter;

namespace Debugging
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            // Create document by specifying paper size and orientation,
            // and default language.
            var doc = new RtfDocument(PaperSize.A4, PaperOrientation.Landscape, Lcid.English);

            // Create fonts and colors for later use
            var times = doc.createFont("Times New Roman");
            var courier = doc.createFont("Courier New");
            var red = doc.createColor(new RtfColor("ff0000"));
            var blue = doc.createColor(new RtfColor(0, 0, 255));
            var white = doc.createColor(new RtfColor(255, 255, 255));
            var colourTableHeader = doc.createColor(new RtfColor("76923C"));
            var colourTableRow = doc.createColor(new RtfColor("D6E3BC"));
            var colourTableRowAlt = doc.createColor(new RtfColor("FFFFFF"));

            // Don't instantiate RtfTable, RtfParagraph, and RtfImage objects by using
            // ``new'' keyword. Instead, use add* method in objects derived from
            // RtfBlockList class. (See Demos.)
            RtfTable table;
            RtfParagraph par;
            RtfImage img;
            // Don't instantiate RtfCharFormat by using ``new'' keyword, either.
            // An addCharFormat method are provided by RtfParagraph objects.
            RtfCharFormat fmt;


            // ==========================================================================
            // Demo 1: Font Setting
            // ==========================================================================
            // If you want to use Latin characters only, it is as simple as assigning
            // ``Font'' property of RtfCharFormat objects. If you want to render Far East
            // characters with some font, and Latin characters with another, you may
            // assign the Far East font to ``Font'' property and the Latin font to
            // ``AnsiFont'' property.
            par = doc.addParagraph();
            par.Alignment = Align.Left;
            par.DefaultCharFormat.Font = times;
            par.DefaultCharFormat.AnsiFont = courier;
            par.setText("Testing\n");


            // ==========================================================================
            // Demo 2: Character Formatting
            // ==========================================================================
            par = doc.addParagraph();
            par.DefaultCharFormat.Font = times;
            par.setText("Demo2: Character Formatting");
            // Besides setting default character formats of a paragraph, you can specify
            // a range of characters to which formatting is applied. For convenience,
            // let's call it range formatting. The following section sets formatting
            // for the 4th, 5th, ..., 8th characters in the paragraph. (Note: the first
            // character has an index of 0)
            fmt = par.addCharFormat(4, 8);
            fmt.FgColor = blue;
            fmt.BgColor = red;
            fmt.FontSize = 18;
            // Sets another range formatting. Note that when range formatting overlaps,
            // the latter formatting will overwrite the former ones. In the following,
            // formatting for the 8th chacacter is overwritten.
            fmt = par.addCharFormat(8, 10);
            fmt.FontStyle.addStyle(FontStyleFlag.Bold);
            fmt.FontStyle.addStyle(FontStyleFlag.Underline);
            fmt.Font = courier;


            // ==========================================================================
            // Demo 3: Footnote
            // ==========================================================================
            par = doc.addParagraph();
            par.setText("Demo3: Footnote");
            // In this example, the footnote is inserted just after the 7th character in
            // the paragraph.
            par.addFootnote(7).addParagraph().setText("Footnote details here.");


            // ==========================================================================
            // Demo 4: Header and Footer
            // ==========================================================================
            // You may use ``Header'' and ``Footer'' properties of RtfDocument objects to
            // specify information to be displayed in the header and footer of every page,
            // respectively.
            par = doc.Footer.addParagraph();
            par.setText("Demo4: Page: / Date: Time:");
            par.Alignment = Align.Center;
            par.DefaultCharFormat.FontSize = 15;
            // You may insert control words, including page number, total pages, date and
            // time, into the header and/or the footer.
            par.addControlWord(12, RtfFieldControlWord.FieldType.Page);
            par.addControlWord(13, RtfFieldControlWord.FieldType.NumPages);
            par.addControlWord(19, RtfFieldControlWord.FieldType.Date);
            par.addControlWord(25, RtfFieldControlWord.FieldType.Time);
            // Here we also add some text in header.
            par = doc.Header.addParagraph();
            par.setText("Demo4: Header");


            // ==========================================================================
            // Demo 5: Image
            // ==========================================================================
            img = doc.addImage("../../demo5.jpg", ImageFileType.Jpg);
            // You may set the width only, and let the height be automatically adjusted
            // to keep aspect ratio.
            img.Width = 130;
            // Place the image on a new page. The ``StartNewPage'' property is also supported
            // by paragraphs and tables.
            //img.StartNewPage = true;
            img.StartNewPara = true;


            // ==========================================================================
            // demo 6: 表格
            // ==========================================================================
            // Please be careful when dealing with tables, as most crashes come from them.
            // If you follow steps below, the resulting RTF is not likely to crash your
            // MS Word.
            // 
            // Step 1. Plan and draw the table you want on a scratch paper.
            // Step 2. Start with a MxN regular table.
            table = doc.addTable(5, 4, 415.2f, 12);
            table.Margins[Direction.Bottom] = 20;
            table.setInnerBorder(BorderStyle.Dotted, 1f);
            table.setOuterBorder(BorderStyle.Single, 2f);

            table.HeaderBackgroundColour = colourTableHeader;
            table.RowBackgroundColour = colourTableRow;
            table.RowAltBackgroundColour = colourTableRowAlt;


            // Step 3. (Optional) Set text alignment for each cell, row height, column width,
            //			border style, etc.
            for (var i = 0; i < table.RowCount; i++) {
                for(var j = 0; j < table.ColCount; j++) {
                    table.cell(i, j).addParagraph().setText("CELL " + i.ToString() + "," + j.ToString());
                }
            }

            // Step 4. Merge cells so that the resulting table would look like the one you drew
            //			on paper. One cell cannot be merged twice. In this way, we can construct
            //			almost all kinds of tables we need.
            table.merge(1, 0, 3, 1);
            // Step 5. You may start inserting content for each cell. Actually, it is adviced
            //			that the only thing you do after merging cell is inserting content.
            table.cell(4, 3).BackgroundColour = red;
            table.cell(4, 3).addParagraph().setText("Demo6: Table");


            // ==========================================================================
            // Demo 7: ``Two in one'' format
            // ==========================================================================
            // This format is provisioned for Far East languages. This demo uses Traditional
            // Chinese as an example.
            par = doc.addParagraph();
            par.setText("Demo7: aaa並排文字aaa");
            fmt = par.addCharFormat(10, 13);
            fmt.TwoInOneStyle = TwoInOneStyle.Braces;
            fmt.FontSize = 16;


            // ==========================================================================
            // Demo 7.1: Hyperlink
            // ==========================================================================
            par = doc.addParagraph();
            par.setText("Demo 7.1: Hyperlink to target (Demo9)");
            fmt = par.addCharFormat(10, 18);
            fmt.LocalHyperlink = "target";
            fmt.FgColor = blue;


            // ==========================================================================
            // Demo 8: New page
            // ==========================================================================
            par = doc.addParagraph();
            par.StartNewPage = true;
            par.setText("Demo8: New page");


            // ==========================================================================
            // Demo 9: Set bookmark
            // ==========================================================================
            par = doc.addParagraph();
            par.setText("Demo9: Set bookmark");
            fmt = par.addCharFormat(0, 18);
            fmt.Bookmark = "target";


            // ==========================================================================
            // Save
            // ==========================================================================
            // You may also retrieve RTF code string by calling to render() method of
            // RtfDocument objects.
            doc.save("Demo.rtf");


            // ==========================================================================
            // Open the RTF file we just saved
            // ==========================================================================
            var p = new Process {StartInfo = {FileName = "Demo.rtf"}};
            p.Start();
        }
    }
}
