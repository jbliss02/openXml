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
    public class SalesReportBuilder
    {

        const string drawingTemplate = @"c:\temp\drawingTemplate.xml";
        const string headerImageFile = @"c:\temp\headerimage.gif";
        const string stylesXmlFile = @"c:\temp\styles.xml";
        const string wordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const string wordPrefix = "w";

        private string _documentName;
        private string _imagePartRID;

        public SalesReportBuilder()
        {

            //Generate unique filename
            // _documentName = @"~/reports/AdventureWorks" + DateTime.Now.ToFileTime() + ".docx";
            _documentName = @"c:\temp\testmedoc.docx";
            Console.WriteLine(CreateDocument());
          //  openAndAddHeader();
            Console.ReadLine();
            string s = "j";
        }

        #region Create Package and Parts
        /// <summary>
        /// 1. Create a new package as a Word document.
        /// 2. Add a style.xml part.
        /// 3. Add an embedded image part.
        /// 4. Create the document.xml part content. 
        /// </summary>
        /// <returns>File path location or error message</returns>
        public string CreateDocument()
        {
            //try
            //{
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(_documentName, WordprocessingDocumentType.Document))
                {

                    // Set the content of the document so that Word can open it.
                    MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();

                    //header
                    //mainPart.DeleteParts(mainPart.HeaderParts);
                    //HeaderPart headerPart = wordDoc.AddNewPart<HeaderPart>();
                    //string headerPartId = mainPart.GetIdOfPart(headerPart);

                    //GenerateHeaderPartContent(headerPart);

                    // Get SectionProperties and Replace HeaderReference and FooterRefernce with new Id
                    //IEnumerable<SectionProperties> sections = mainPart.Document.Body.Elements<SectionProperties>();

                    //foreach (var section in sections)
                    //{
                    //    // Delete existing references to headers and footers
                    //    section.RemoveAllChildren<HeaderReference>();
                    //    section.RemoveAllChildren<FooterReference>();

                    //    // Create the new header and footer reference node
                    //    section.PrependChild<HeaderReference>(new HeaderReference() { Id = headerPartId });
                    //    section.PrependChild<FooterReference>(new FooterReference() { Id = footerPartId });
                    //}

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

                    // Create an image part and add it to the document.
                    ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Gif);
                    string imageFileName = headerImageFile;
                    using (FileStream stream = new FileStream(imageFileName, FileMode.Open))
                    {
                        imagePart.FeedData(stream);
                    }

                    // Get the reference ID for the image added to the package.
                    // You will use the image part reference ID to insert the
                    // image to the document.xml file.
                    _imagePartRID = mainPart.GetIdOfPart(imagePart);

                    // Create document.xml content.
                    SetMainDocumentContent(mainPart);

                    
                }

               
                return (_documentName);
            //}
            //catch (Exception ex)
            //{
            //    return (ex.Message);

            //}
        }

        public void openAndAddHeader()
        {
            using (WordprocessingDocument document = WordprocessingDocument.Open(_documentName, true))
            {
                // Get the main document part
                MainDocumentPart mainDocumentPart = document.MainDocumentPart;

                // Delete the existing header and footer parts
                mainDocumentPart.DeleteParts(mainDocumentPart.HeaderParts);
                //mainDocumentPart.DeleteParts(mainDocumentPart.FooterParts);

                // Create a new header and footer part
                HeaderPart headerPart = mainDocumentPart.AddNewPart<HeaderPart>();
                //FooterPart footerPart = mainDocumentPart.AddNewPart<FooterPart>();

                // Get Id of the headerPart and footer parts
                string headerPartId = mainDocumentPart.GetIdOfPart(headerPart);
                //string footerPartId = mainDocumentPart.GetIdOfPart(footerPart);

                GenerateHeaderPartContent(headerPart);

                //GenerateFooterPartContent(footerPart);

                //// Get SectionProperties and Replace HeaderReference and FooterRefernce with new Id

                //for(int i = 0; i < sections.)

                var sections = mainDocumentPart.Document.Body.Elements<SectionProperties>();

                foreach (var section in sections)
                {
                    // Delete existing references to headers and footers
                    section.RemoveAllChildren<HeaderReference>();
                    section.RemoveAllChildren<FooterReference>();

                    // Create the new header and footer reference node
                    section.PrependChild<HeaderReference>(new HeaderReference() { Id = headerPartId });
                    //.PrependChild<FooterReference>(new FooterReference() { Id = footerPartId });
                }
            }
        }

        public void GenerateHeaderPartContent(HeaderPart part)
        {
            Header header1 = new Header() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 wp14" } };
            header1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            header1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            header1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            header1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            header1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            header1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            header1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            header1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            header1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            header1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            header1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            header1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            header1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            header1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            header1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "00164C17", RsidRunAdditionDefault = "00164C17" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Header" };

            paragraphProperties1.Append(paragraphStyleId1);

            Run run1 = new Run();
            Text text1 = new Text();
            text1.Text = "Header McSpedder";

            run1.Append(text1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);

            header1.Append(paragraph1);

            part.Header = header1;
        }

        /// <summary>
        /// Set content of MainDocumentPart. 
        /// </summary>
        /// <param name="part">MainDocumentPart</param>
        public void SetMainDocumentContent(MainDocumentPart part)
        {
            using (Stream stream = part.GetStream())
            {
                CreateWordProcessingML(stream);
            }
        }


        /// <summary>
        /// Generate WordprocessingML for Sales Report.
        /// The resulting XML will be saved as document.xml.
        /// </summary>
        /// <param name="stream">MainDocumentPart stream</param>
        public void CreateWordProcessingML(Stream stream)
        {

            // Get sales person data from AdventureWorks database
            // You will write this data to the document.xml file.
            //AdventureWorksSalesData salesData = new AdventureWorksSalesData();
            //StringDictionary SalesPerson = salesData.GetSalesPersonData(_salesPersonID);

            // Create an XmlWriter using UTF8 encoding.
            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Encoding = Encoding.UTF8;
            settings.Indent = true;

            // This file represents the WordprocessingML of the Sales Report.
            XmlWriter writer = XmlWriter.Create(stream, settings);
            //try
            //{
                writer.WriteStartDocument(true);
                writer.WriteComment("This file represents the WordProcessingML of our Sales Report");
                writer.WriteStartElement(wordPrefix, "document", wordNamespace);
                writer.WriteStartElement(wordPrefix, "body", wordNamespace);

                WriteHeaderImage(writer);
                WriteDocumentTitle(writer, "Some bullshit title");
                //WriteDocumentContactInfo(writer, SalesPerson["FullName"], SalesPerson["Phone"], SalesPerson["Email"]);
                //WriteSalesSummaryInfo(writer, SalesPerson["SalesYTD"], SalesPerson["SalesQuota"]);
                //WriteTerritoriesTable(writer, SalesPerson["TerritoryName"]);

                writer.WriteEndElement(); //body
                writer.WriteEndElement(); //document
            //}
            //catch (Exception e)
            //{
            //    throw;
            //}
            //finally
            //{
                //Write the XML to file and close the writer.
                writer.Flush();
                writer.Close();
            //}
            return;
        }

        #endregion

        #region Formatting Methods
        /// <summary>
        /// Write the title paragraph properties to the WordprocessingML document.
        /// </summary>
        /// <param name="writer">The XmlWriter to write the properties to.</param>
        
        private void WriteParaHeader(XmlWriter writer){

        }
        
        
        private void WriteTitleParagraphProperties(XmlWriter writer)
        {
            // Create the paragraph properties element.
            // </w:pPr>
            writer.WriteStartElement(wordPrefix, "pPr",
               wordNamespace);

            // Create the bottom border.
            //   <w:pBdr>
            //     <w:bottom w:val=”single” w:sz=”4” 
            //               w:space=”1” w:color=”auto” />
            //   </w:pBdr>
            writer.WriteStartElement(wordPrefix, "pBdr", wordNamespace);
            writer.WriteStartElement(wordPrefix, "bottom", wordNamespace);
            writer.WriteAttributeString(wordPrefix, "val", wordNamespace, "single");
            writer.WriteAttributeString(wordPrefix, "sz", wordNamespace, "4");
            writer.WriteAttributeString(wordPrefix, "space", wordNamespace, "1");
            writer.WriteAttributeString(wordPrefix, "color", wordNamespace, "blue");
            writer.WriteEndElement();
            writer.WriteEndElement();

            // Define the spacing for the paragraph.
            //   <w:spacing w:line=”240” w:lineRule=”auto” />
            writer.WriteStartElement(wordPrefix, "spacing", wordNamespace);
            writer.WriteAttributeString(wordPrefix, "line", wordNamespace, "240");
            writer.WriteAttributeString(wordPrefix, "lineRule", wordNamespace, "auto");
            writer.WriteEndElement();

            // Close the properties element.
            // </w:pPr>
            writer.WriteEndElement();
        }

        /// <summary>
        /// Write the title run properties to the WordprocessingML document.
        /// </summary>
        /// <param name="writer">The XmlWriter to write the properties to.</param>
        private void WriteTitleRunProperties(XmlWriter writer)
        {
            // Create the run properties element.
            // <w:rPr>
            writer.WriteStartElement(wordPrefix, "rPr", wordNamespace);

            // Set up the spacing.
            //   <w:spacing w:val=”5” />
            writer.WriteStartElement(wordPrefix, "spacing", wordNamespace);
            writer.WriteAttributeString(wordPrefix, "val", wordNamespace, "5");
            writer.WriteEndElement();

            // Define the size.
            //   <w:sz w:val=”52” />
            writer.WriteStartElement(wordPrefix, "sz", wordNamespace);
            writer.WriteAttributeString(wordPrefix, "val", wordNamespace, "52");
            writer.WriteEndElement();

            // Close the properties element.
            // </w:rPr>
            writer.WriteEndElement();
        }

        /// <summary>
        /// Write the subtitle paragraph properties to the WordprocessingML document.
        /// </summary>
        /// <param name="writer">The XmlWriter to write the properties to.</param>
        private void WriteSubtitleParagraphProperties(XmlWriter writer)
        {
            // Create the paragraph properties element.
            // <w:pPr>
            writer.WriteStartElement(wordPrefix, "pPr", wordNamespace);

            // Define the spacing for the paragraph.
            //   <w:spacing w:before=”200” w:after=”0” />
            writer.WriteStartElement(wordPrefix, "spacing", wordNamespace);
            writer.WriteAttributeString(wordPrefix, "before", wordNamespace, "200");
            writer.WriteAttributeString(wordPrefix, "after", wordNamespace, "0");
            writer.WriteEndElement();

            // Close the properties element.
            // </w:pPr>
            writer.WriteEndElement();
        }

        /// <summary>
        /// Write the subtitle run properties to the WordprocessingML document.
        /// </summary>
        /// <param name="writer">The XmlWriter to write the properties to.</param>
        private void WriteSubtitleRunProperties(XmlWriter writer)
        {
            // Create the run properties element.
            // <w:rPr>
            writer.WriteStartElement(wordPrefix, "rPr", wordNamespace);

            // setup as bold
            //   <w:b />
            writer.WriteElementString(wordPrefix, "b", wordNamespace, null);

            // Define the size.
            //   <sz w:val=”26” />
            writer.WriteStartElement(wordPrefix, "sz", wordNamespace);
            writer.WriteAttributeString(wordPrefix, "val", wordNamespace, "26");
            writer.WriteEndElement();

            // Close the properties element.
            // </w:rPr>
            writer.WriteEndElement();
        }

        /// <summary>
        /// Write the bold run property to the WordprocessingML document.
        /// </summary>
        /// <param name="writer">The XmlWriter to write the properties to.</param>
        private void WriteBoldRunProperties(XmlWriter writer)
        {
            // Create the run properties element.
            // <w:rPr>
            //   <w:b />
            // </w:rPr>
            writer.WriteStartElement(wordPrefix, "rPr", wordNamespace);
            writer.WriteElementString(wordPrefix, "b", wordNamespace, null);
            writer.WriteEndElement();
        }
        #endregion

        #region Styles Methods
    /// <summary>
    /// Write the style property to the WordprocessingML paragraph element.
    /// </summary>
    /// <param name="writer">The XmlWriter to write the style to.</param>
    public void ApplyParagraphStyle(XmlWriter writer, string styleId) {
        // Apply the style in the paragraph properties.
        // <w:pPr>
        //   <w:pStyle w:val=”MyTitle” />
        // </w:pPr>
        writer.WriteStartElement(wordPrefix, "pPr", wordNamespace);
        writer.WriteStartElement(wordPrefix, "pStyle", wordNamespace);
        writer.WriteAttributeString(wordPrefix, "val", wordNamespace, styleId);
        writer.WriteEndElement();
        writer.WriteEndElement();
    }

    /// <summary>
    /// Write the style property to the WordprocessingML table element.
    /// </summary>
    /// <param name="writer">The XmlWriter to write the style to.</param>
    public void ApplyTableStyle(XmlWriter writer, string styleId) {
        // Apply the style in the table properties.
        // <w:tblPr>
        //   <w:tblStyle w:val="MyTableStyle" />
        //   <w:tblW w:w="0" w:type="auto" /> 
        //   <w:tblLook w:val="04A0" /> 
        // </w:tblPr>
        writer.WriteStartElement(wordPrefix, "tblPr", wordNamespace);
        writer.WriteStartElement(wordPrefix, "tblStyle", wordNamespace);
        writer.WriteAttributeString(wordPrefix, "val", wordNamespace, styleId);
        writer.WriteEndElement();
        writer.WriteStartElement(wordPrefix, "tblW", wordNamespace);
        writer.WriteAttributeString(wordPrefix, "w", wordNamespace, "0");
        writer.WriteAttributeString(wordPrefix, "type", wordNamespace, "auto");
        writer.WriteEndElement();
        writer.WriteStartElement(wordPrefix, "tblLook", wordNamespace);
        writer.WriteAttributeString(wordPrefix, "val", wordNamespace, "04A0");
        writer.WriteEndElement();
        writer.WriteEndElement();
    }
    #endregion

        #region document.xml writing methods
    /// <summary>
    /// Writes an image within a paragraph
    /// into the WordprocessingML.
    /// </summary>
    /// <param name="writer">The XmlWriter to write the image to.</param>
        private void WriteHeaderImage(XmlWriter writer) {
            // Load the drawing template into an XML document.
            XmlDocument drawingXml = new XmlDocument();
            string drawingXmlFile = drawingTemplate;
            drawingXml.Load(drawingXmlFile);

            // Load the drawing template into an XML document and pass the reference ID parameter.
            drawingXml.LoadXml(string.Format(drawingXml.InnerXml, _imagePartRID));

            // Write the wrapping paragraph and the drawing fragment.
            writer.WriteStartElement(wordPrefix, "p", wordNamespace);
            writer.WriteStartElement(wordPrefix, "r", wordNamespace);
            drawingXml.DocumentElement.WriteContentTo(writer);
            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        private void WriteDocumentTitle(XmlWriter writer, string title)
        {
            // Create the title.
            // <w:p>
            //   <w:r>
            //     <w:t>Sales Report - Employee Name</w:t>
            //   </w:r>
            // </w:p>    

            writer.WriteStartElement(wordPrefix, "p", wordNamespace);
            WriteTitleParagraphProperties(writer);
            writer.WriteStartElement(wordPrefix, "r", wordNamespace);
            WriteTitleRunProperties(writer);
            writer.WriteElementString(wordPrefix, "t", wordNamespace, "Sales Report - " + title);
            writer.WriteEndElement();
            writer.WriteEndElement();
        }



        #endregion
    }
}
