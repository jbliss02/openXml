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
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(_documentName,WordprocessingDocumentType.Document))
            {
                // Add a new main document part. 
                mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new Document();
                
                //Create Body (this element contains other elements that we want to include 
                body = new Body();

                string s = "kld";
                //add styles
                addExistingStyles();
                //addNewStyle(mainPart);

                pleasework();

                body.Append(returnTextInput());
                body.Append(returnTextInput());

                addSectionHeader("Users and user limits");
                addSectionText("This is where the random text comes, blah blah blah");
                addSectionText("Another line.");
                addSectionBreak(SectionMarkValues.Continuous);
                addSectionText("Another line.");
                addTable();
                addTextBox();
              
                ////table
                Table table = new Table(new TableRow(new TableCell(new Paragraph(new Run(new Text("Hello World!"))))));
                body.Append(table);
                mainPart.Document.Append(body);

                setMargins();
                addHeader("THE HEADING!");
                addFooter("txt");
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

            paragraph.AppendChild(new Paragraph(new ParagraphProperties(new SectionProperties(new SectionType() { Val = SectionMarkValues.Continuous }))));

            body.Append(paragraph);

            // Save changes to the main document part. 
        }

        private void addTextBox()
        {

            FormFieldData fd = new FormFieldData();
            TextInput txt = new TextInput();
        
            fd.AppendChild(txt);

            Paragraph paragraph = new Paragraph();
            Run run_paragraph = new Run();
            run_paragraph.Append(fd);
            paragraph.Append(run_paragraph);
            body.Append(paragraph);

;

            return;

            BookmarkStart bkmStart = new BookmarkStart() { Name = "Table1", Id = "Table1" };
            BookmarkEnd bkmEnd = new BookmarkEnd() { Id = "Table1" };
            body.Append(bkmStart);
        }

        private void addTable()
        {
            Table table = new Table();
           
            TableProperties tblPr = new TableProperties();
            tblPr.TableStyle = new TableStyle();
            tblPr.TableStyle.Val = "ocSummary";

            TableWidth tableWidth = new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct };
            TableBorders tblBorders = new TableBorders();

            #region ramblings
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
            #endregion
            tblPr.Append(tblBorders, tableWidth);
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
                    //tr.Append(new TableCell(new Paragraph(new Run(new Text((i * j).ToString())))));
                    tr.Append(new TableCell(returnTextInput()));
                }
                table.Append(tr);
            }


            //appending table to body
            body.Append(table);



        }

        private void addFooter(string text)
        {
            //footer
            FooterPart newFtPart = mainPart.AddNewPart<FooterPart>();
            string ft_ID = mainPart.GetIdOfPart(newFtPart);

            returnFooter(text).Save(newFtPart);
            foreach (SectionProperties sectProperties in mainPart.Document.Descendants<SectionProperties>())
            {
                FooterReference newFtReference =
                 new FooterReference() { Id = ft_ID, Type = HeaderFooterValues.Default };
                sectProperties.Append(newFtReference);
            }

        }//addFooter()

        private void addHeader(string text)
        {
            HeaderPart newFtPart = mainPart.AddNewPart<HeaderPart>();
            string ft_ID = mainPart.GetIdOfPart(newFtPart);

            returnHeader(text).Save(newFtPart);
            foreach (SectionProperties sectProperties in mainPart.Document.Descendants<SectionProperties>())
            {
                HeaderReference newFtReference =
                 new HeaderReference() { Id = ft_ID, Type = HeaderFooterValues.Default };
                sectProperties.Append(newFtReference);
            }


        }//addHeader()

        private void addSectionBreak(SectionMarkValues breakType)
        {
            //has to be appended to a paragraph, otherwise is a page break

            //body.AppendChild(new Paragraph(new Run(new SectionProperties( new SectionType() { Val = SectionMarkValues.Continuous }))));

            body.AppendChild(new Paragraph(new ParagraphProperties(new SectionProperties(new SectionType() { Val = SectionMarkValues.Continuous }))));

            return;


            Paragraph p = new Paragraph();

            ParagraphProperties p_prop = new ParagraphProperties();

            SectionProperties sectionProperties1 = new SectionProperties();
            SectionType sectionType1 = new SectionType() { Val = SectionMarkValues.Continuous };

            sectionProperties1.Append(sectionType1);

            p_prop.Append(sectionProperties1);

            p.Append(p_prop);

            Run run_paragraph = new Run();
            Text text_paragraph = new Text("-");
            run_paragraph.Append(text_paragraph);
            p.Append(run_paragraph);


            body.Append(p);
        }

        private void addCustomerBox()
        {
            XmlDocument custXml = new XmlDocument();
            custXml.Load(@"c:\temp\customerTemplate.xml");

            using (Stream outputStream = mainPart.GetStream())
            {
                using (StreamWriter ts = new StreamWriter(outputStream))
                {
                    ts.Write(custXml.InnerXml);
                }
            }


        }

        private Paragraph returnTextInput()
        {
            Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "00A079F1", RsidRunAdditionDefault = "008E35FE" };

            Run run1 = new Run();

            FieldChar fieldChar1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

            FormFieldData formFieldData1 = new FormFieldData();
            FormFieldName formFieldName1 = new FormFieldName() { Val = "Text35" };
            Enabled enabled1 = new Enabled();
            CalculateOnExit calculateOnExit1 = new CalculateOnExit() { Val = false };
            TextInput textInput1 = new TextInput();

            formFieldData1.Append(formFieldName1);
            formFieldData1.Append(enabled1);
            formFieldData1.Append(calculateOnExit1);
            formFieldData1.Append(textInput1);

            fieldChar1.Append(formFieldData1);

            run1.Append(fieldChar1);

            Run run2 = new Run();
            FieldCode fieldCode1 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
            fieldCode1.Text = " FORMTEXT ";

            run2.Append(fieldCode1);

            Run run3 = new Run();
            FieldChar fieldChar2 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

            run3.Append(fieldChar2);

            Run run4 = new Run();

            RunProperties runProperties1 = new RunProperties();
            NoProof noProof1 = new NoProof();

            runProperties1.Append(noProof1);
            Text text1 = new Text();
            text1.Text = " ";

            run4.Append(runProperties1);
            run4.Append(text1);

            Run run5 = new Run();

            RunProperties runProperties2 = new RunProperties();
            NoProof noProof2 = new NoProof();

            runProperties2.Append(noProof2);
            Text text2 = new Text();
            text2.Text = " ";

            run5.Append(runProperties2);
            run5.Append(text2);

            Run run6 = new Run();

            RunProperties runProperties3 = new RunProperties();
            NoProof noProof3 = new NoProof();

            runProperties3.Append(noProof3);
            Text text3 = new Text();
            text3.Text = " ";

            run6.Append(runProperties3);
            run6.Append(text3);

            Run run7 = new Run();

            RunProperties runProperties4 = new RunProperties();
            NoProof noProof4 = new NoProof();

            runProperties4.Append(noProof4);
            Text text4 = new Text();
            text4.Text = " ";

            run7.Append(runProperties4);
            run7.Append(text4);

            Run run8 = new Run();

            RunProperties runProperties5 = new RunProperties();
            NoProof noProof5 = new NoProof();

            runProperties5.Append(noProof5);
            Text text5 = new Text();
            text5.Text = " ";

            run8.Append(runProperties5);
            run8.Append(text5);

            Run run9 = new Run();
            FieldChar fieldChar3 = new FieldChar() { FieldCharType = FieldCharValues.End };

            run9.Append(fieldChar3);
            BookmarkStart bookmarkStart1 = new BookmarkStart() { Name = "_GoBack", Id = "0" };
            BookmarkEnd bookmarkEnd1 = new BookmarkEnd() { Id = "0" };

            paragraph1.Append(run1);
            paragraph1.Append(run2);
            paragraph1.Append(run3);
            paragraph1.Append(run4);
            paragraph1.Append(run5);
            paragraph1.Append(run6);
            paragraph1.Append(run7);
            paragraph1.Append(run8);
            paragraph1.Append(run9);
            paragraph1.Append(bookmarkStart1);
            paragraph1.Append(bookmarkEnd1);

            //body.Append(paragraph1);
            return paragraph1;
        }

        private void pleasework()
        {
            Paragraph paragraph1 = new Paragraph();

            SimpleField simpleField = new SimpleField(); simpleField.Instruction = "FORMTEXT";

            FormFieldData formFieldData = new FormFieldData();

            formFieldData.Append(new FormFieldName() { Val = "MyField" });

            formFieldData.Append(new Enabled() { Val = OnOffValue.FromBoolean(false) });

            TextInput textInput = new TextInput();

            textInput.Append(new MaxLength() { Val = 4 });
            formFieldData.Append(textInput);
            simpleField.Append(formFieldData);
            paragraph1.Append(simpleField);
            body.Append(paragraph1);

        }

        private Footer returnFooter(string text)
        {
            var element =
              new Footer(
                new Paragraph(
                  new ParagraphProperties(
                    new ParagraphStyleId() { Val = "Footer" }),
                  new Run(
                    new Text(text))
                )
              );

            return element;
        }

        private Header returnHeader(string text)
        {
            var element =
              new Header(
                new Paragraph(
                  new ParagraphProperties(
                    new ParagraphStyleId() { Val = "Header" }),
                  new Run(
                    new Text(text))
                )
  );

            return element;
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
