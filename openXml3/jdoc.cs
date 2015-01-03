using System;
using System.Data;
using System.Collections;
using System.Collections.Specialized;
using System.Configuration;
using System.IO;
using System.IO.Packaging;
using System.Text;
using System.Xml;
using System.Web;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
namespace openXml3
{
    public class jdoc
    {
        private string _documentName;
        private string _imagePartRID;
        const string stylesXmlFile = @"c:\temp\styles.xml";
        private MainDocumentPart mainPart;
        private Body body;
        public jdoc()
        {
            _documentName = @"c:\temp\testmedoc.docx";
            
        }


        public void createDoc()
        {
            // Create a Wordprocessing document. 
            using (WordprocessingDocument myDoc = WordprocessingDocument.Create(_documentName,WordprocessingDocumentType.Document))
            {
                // Add a new main document part. 
                mainPart = myDoc.AddMainDocumentPart();
                mainPart.Document = new Document();
                
                //Create Body (this element contains other elements that we want to include 
                body = new Body();

                //add styles
                addExistingStyles();
                //addNewStyle(mainPart);

                addSectionHeader("Users and user limits");
                addSectionText("This is where the random text comes, blah blah blah");
                addTable();

                //table
                Table table = new Table(new TableRow(new TableCell(new Paragraph(new Run(new Text("Hello World!"))))));
                body.Append(table);

                setMargins();

                mainPart.Document.Save();
            } 
        }//createDoc

        private void addSectionHeader(string sectionName)
        {
            Paragraph heading = new Paragraph();
            Run heading_run = new Run();
            Text heading_text = new Text(sectionName);
            ParagraphProperties heading_pPr = new ParagraphProperties();
            // we set the style
            heading_pPr.ParagraphStyleId = new ParagraphStyleId() { Val = "OneContractcalcheader" };
            heading.Append(heading_pPr);
            heading_run.Append(heading_text);
            heading.Append(heading_run);




            mainPart.Document.Append(heading);
        }

        private void addSectionText(string text)
        {
            //Create paragraph 
            Paragraph paragraph = new Paragraph();
            Run run_paragraph = new Run();
            Text text_paragraph = new Text(text);
            run_paragraph.Append(text_paragraph);
            paragraph.Append(run_paragraph);
            body.Append(paragraph);
            mainPart.Document.Append(body);
            // Save changes to the main document part. 
        }

        private void addTable()
        {
            Table table = new Table();
           
            TableProperties tblPr = new TableProperties();
            tblPr.TableStyle = new TableStyle();
            tblPr.TableStyle.Val = "OneContractNoBorder";


            TableBorders tblBorders = new TableBorders();

            //tblBorders.TopBorder = new TopBorder();
            //tblBorders.TopBorder.Val = new EnumValue<BorderValues>(BorderValues.Single);
            //tblBorders.BottomBorder = new BottomBorder();
            //tblBorders.BottomBorder.Val = new EnumValue<BorderValues>(BorderValues.Single);
            //tblBorders.LeftBorder = new LeftBorder();
            //tblBorders.LeftBorder.Val = new EnumValue<BorderValues>(BorderValues.Single);
            //tblBorders.RightBorder = new RightBorder();
            //tblBorders.RightBorder.Val = new EnumValue<BorderValues>(BorderValues.Single);
            //tblBorders.InsideHorizontalBorder = new InsideHorizontalBorder();
            //tblBorders.InsideHorizontalBorder.Val = BorderValues.Single;
            //tblBorders.InsideVerticalBorder = new InsideVerticalBorder();
            //tblBorders.InsideVerticalBorder.Val = BorderValues.Single;
            tblPr.Append(tblBorders);
            table.Append(tblPr);
            TableRow tr;
            TableCell tc;
            //first row - title
            tr = new TableRow();
            tc = new TableCell(new Paragraph(new Run(
                               new Text("Multiplication table"))));
            TableCellProperties tcp = new TableCellProperties();
            GridSpan gridSpan = new GridSpan();
            gridSpan.Val = 11;
            tcp.Append(gridSpan);
            tc.Append(tcp);
            tr.Append(tc);
            table.Append(tr);
            //second row 
            tr = new TableRow();
            tc = new TableCell();
            tc.Append(new Paragraph(new Run(new Text("*"))));
            tr.Append(tc);
            for (int i = 1; i <= 10; i++)
            {
                tr.Append(new TableCell(new Paragraph(new Run(new Text(i.ToString())))));
            }
            table.Append(tr);
            for (int i = 1; i <= 10; i++)
            {
                tr = new TableRow();
                tr.Append(new TableCell(new Paragraph(new Run(new Text(i.ToString())))));
                for (int j = 1; j <= 10; j++)
                {
                    tr.Append(new TableCell(new Paragraph(new Run(new Text((i * j).ToString())))));
                }
                table.Append(tr);
            }


            //appending table to body
            body.Append(table);



        }

        private void setMargins()
        {
            SectionProperties sectionProps = new SectionProperties();
            PageMargin pageMargin = new PageMargin() { Top = 1008, Right = (UInt32Value)1008U, Bottom = 1008, Left = (UInt32Value)1008U, Header = (UInt32Value)720U, Footer = (UInt32Value)720U, Gutter = (UInt32Value)0U };
            sectionProps.Append(pageMargin);
            mainPart.Document.Body.Append(sectionProps);
        }

        private void addExistingStyles()
        {
            // Create a style part and add it to the document.
            XmlDocument stylesXml = new XmlDocument();
            stylesXml.Load(stylesXmlFile);

            StyleDefinitionsPart stylePart = mainPart.AddNewPart<StyleDefinitionsPart>();

            //  Copy the style.xml content into the new part....
            using (Stream outputStream = stylePart.GetStream())
            {
                using (StreamWriter ts = new StreamWriter(outputStream))
                {
                    ts.Write(stylesXml.InnerXml);
                }
            }

        }

        private void addNewStyle()
        {
            StyleDefinitionsPart stylePart = mainPart.AddNewPart<StyleDefinitionsPart>();

            RunProperties rPr = new RunProperties();
            Color color = new Color() { Val = "FF0000" }; // the color is red
            RunFonts rFont = new RunFonts();
            rFont.Ascii = "Arial"; // the font is Arial
            rPr.Append(color);
            rPr.Append(rFont);
            rPr.Append(new Bold()); // it is Bold
            rPr.Append(new FontSize() { Val = "28" }); //font size (in 1/72 of an inch)

            //creation of a style
            Style style = new Style();
            style.StyleId = "MyHeading1"; //this is the ID of the style
            style.Append(new Name() { Val = "My Heading 1" }); //this is name
            // our style based on Normal style
            style.Append(new BasedOn() { Val = "Heading1" });
            // the next paragraph is Normal type
            style.Append(new NextParagraphStyle() { Val = "Normal" });
            style.Append(rPr);//we are adding properties previously defined

            stylePart.Styles = new Styles();
            stylePart.Styles.Append(style);
            stylePart.Styles.Save(); // we save the style part



        }

    }
}
