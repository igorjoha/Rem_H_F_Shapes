using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace Aspose.Words.Examples
{
    public class ExampletHelloWorld
    {
        static void Main(string[] args)
        {
            // Create a new empty document A
            Document dstDocA = new Document();

            // Inisialize a DocumentBuilder
            DocumentBuilder builder = new DocumentBuilder(dstDocA);

            // Get the new section and its page setup.
            Section currentSection = builder.CurrentSection;
            PageSetup pageSetup = currentSection.PageSetup;



            // Insert text to the destination document A start
            // builder.MoveToDocumentStart();
            //builder.Write("My first aspose doc");

            // Open an existing document B
            Document SrcDocB = new Document("D:\\from_old\\Aspose\\copla.docx");

            // Set the source document to continue straight after the end of the destination document.
            SrcDocB.FirstSection.PageSetup.SectionStart = SectionStart.Continuous;

            // Restart the page numbering on the start of the source document.
            SrcDocB.FirstSection.PageSetup.RestartPageNumbering = true;
            SrcDocB.FirstSection.PageSetup.PageStartingNumber = 1;


            // Find the footer that we want to change and remove footer
            //HeaderFooter primaryFooter = currentSection.HeadersFooters[HeaderFooterType.FooterPrimary];

            foreach (Section section in SrcDocB)
            {
                HeaderFooter footer;
                HeaderFooter header;
                ImageData image;

                footer = section.HeadersFooters[HeaderFooterType.FooterFirst];
                if (footer != null)
                    footer.Remove();

                // Primary footer is the footer used for odd pages.
                footer = section.HeadersFooters[HeaderFooterType.FooterPrimary];
                if (footer != null)
                    footer.Remove();

                footer = section.HeadersFooters[HeaderFooterType.FooterEven];
                if (footer != null)
                    footer.Remove();

                //removing headers. primary - odd, even -even, first -???
                header = section.HeadersFooters[HeaderFooterType.HeaderPrimary];
                if (header != null)
                    header.Remove();


                // Remove all shapes (images, charts, etc.)
                NodeCollection shapes = SrcDocB.GetChildNodes(NodeType.Shape, true);
                shapes.Clear();

            }

            //Add source document B whole doc B to the end of destination document A, preserving document B formatting
            dstDocA.AppendDocument(SrcDocB, ImportFormatMode.KeepSourceFormatting);

            // Save the output as PDF
                dstDocA.Save("D:\\Aspose\\output_smth_new.pdf");

            // Open the source DOCX document.
            //_dataDir = ("D:\\Aspose\\");
            // Document doc = new Document(_dataDir + "input.docx");

            // Save the file to PDF format.
            //SrcDocB.Save(_dataDir + "ouput.pdf", SaveFormat.Docx);


            // Open the source PDF document

            //    Document pdfDocument = new Document(_dataDir + "PDFToDOC.pdf");

            // Save the file into MS document format
            //   pdfDocument.Save(_dataDir + "PDFToDOC_out.doc", SaveFormat.Doc); // .Docx, .Rtf, .WordML, etc.

        }
    }
}
