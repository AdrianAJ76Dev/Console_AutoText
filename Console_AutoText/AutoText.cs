using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Drawing;
using W=DocumentFormat.OpenXml.Wordprocessing;
using System.IO;

namespace Console_AutoText
{
    class AutoText
    {
        /* Keep in mind that these objects are sorta already serialized (they're xml afterall) so I shouldn't worry
         * so much about bringing back an actual object and more conerned about bringing back the xml representation
         */
        private string bodyxml;
        private string imagexml;
        private string embed;

        /* PseudoCode 
         * Pull DocPartBody from Glossary document
         * Use DocPartBody to pull Drawing
         * Use Drawing to pull Blip which is where the image ID is stored
         * I think, IdPartPair is how to reference the document.xml.rels XML document
         * by searching for the id and returning the OpenXMLPart, ImagePart in this case
         * i.e. the Media folder.
         * 
         */

        public static void SimpleWordDoc_Create()
        {
            //string simpledocumentfullname = @"C:\Users\ajones\Documents\Automation\Prototypes\Sole Source Letter\Documents Generated\CB New Simple Document.docx"
            string simpledocumentfullnameHome = @"C:\Users\ajones\Documents\Visual Studio 2015\Operation Kyuzo\Prototypes and Study\Documents Generated\CB New Simple Document.docx";
            //using (WordprocessingDocument NewSimpleDoc = WordprocessingDocument.Create(simpledocumentfullname, WordprocessingDocumentType.Document))
            using (WordprocessingDocument NewSimpleDoc = WordprocessingDocument.Create(simpledocumentfullnameHome, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mdpNewSimpleDoc = NewSimpleDoc.AddMainDocumentPart();
                mdpNewSimpleDoc.Document = new W.Document(new W.Body(new W.Paragraph(new W.Run(new W.Text()))));

                /* Adding image part now */
                mdpNewSimpleDoc.AddImagePart(ImagePartType.Png, "rId7");
                mdpNewSimpleDoc.AddImagePart(ImagePartType.Png, "rId8");
            }
            Console.WriteLine("Done creating the document");
        }

        public static void DisplayImagePartsfromTemplate()
        {
            string TemplateName = @"C:\Users\ajones\Documents\Visual Studio 2015\Operation Kyuzo\Prototypes and Study\Templates\SoleSourceLetter v54.dotx";
            using (WordprocessingDocument wrdTemplate = WordprocessingDocument.Open(TemplateName, false))
            {
                MainDocumentPart mdpTemplate = wrdTemplate.MainDocumentPart;
                Console.WriteLine("Template main document part parts count {0}", mdpTemplate.Parts.Count());
                Console.WriteLine("*************************************************************************");
                foreach (IdPartPair mdpPart in mdpTemplate.Parts)
                {
                    Console.WriteLine("RelationshipID:\t{0}", mdpPart.RelationshipId);
                    Console.WriteLine("OpenXML Part:\t{0}", mdpPart.OpenXmlPart.GetType().Name);
                    Console.WriteLine("Uri:\t\t{0}", mdpPart.OpenXmlPart.Uri);
                    Console.WriteLine();
                }
                Console.WriteLine("*************************************************************************");
                Console.ReadLine();
                /*
                imagecount = mdpTemplate.ImageParts.Count();
                Console.WriteLine("Image Parts count = {0}", imagecount);
                foreach (var image in mdpTemplate.ImageParts)
                {
                    Console.WriteLine("Uri = {0}", image.Uri);
                    Console.WriteLine("RelationshipType = {0}",image.RelationshipType);
                }
                Console.WriteLine("{0}, Done displaying the image files found in the main document part", imagecount);
                Console.WriteLine();
                Console.WriteLine();
                */
                if (wrdTemplate.MainDocumentPart.GetPartsOfType<GlossaryDocumentPart>().Count() > 0)
                {
                    Console.WriteLine("Template Glossary document parts count {0}", mdpTemplate.Parts.Count());
                    Console.WriteLine("*************************************************************************");
                    GlossaryDocumentPart gp = wrdTemplate.MainDocumentPart.GlossaryDocumentPart;
                    /*
                    imagecount = gp.ImageParts.Count();
                    foreach (ImagePart image in gp.ImageParts)
                    {
                        Console.WriteLine("Uri = {0}", image.Uri);
                        Console.WriteLine("RelationshipType = {0}",image.RelationshipType);
                        Console.WriteLine(image);
                        Console.WriteLine();
                    }
                    Console.WriteLine("{0}, Done displaying the image files found in the Glossary Parts document", imagecount);
                    Console.WriteLine();
                    Console.WriteLine(gp.Parts.Count());
                    */

                    foreach (IdPartPair gpart in gp.Parts)
                    {
                        Console.WriteLine("RelationshipID:\t{0}", gpart.RelationshipId);
                        Console.WriteLine("OpenXML Part:\t{0}", gpart.OpenXmlPart.GetType().Name);
                        Console.WriteLine("Uri:\t\t{0}", gpart.OpenXmlPart.Uri);
                        Console.WriteLine();
                    }
                    Console.WriteLine("*************************************************************************");
                }
            }
        }



        //public static void DisplayInnerXML(string AutoTextName)
        //{
        //    string relid=null;
        //    string templatefullname = @"C:\Users\ajones\Documents\Automation\Prototypes\Sole Source Letter\Templates\Sole Source Letter v4.dotx";
        //    using (WordprocessingDocument wrdTemplate = WordprocessingDocument.Open(templatefullname, false))
        //    {
        //        MainDocumentPart mdpTemplate = wrdTemplate.MainDocumentPart;
        //        Document doc = mdpTemplate.Document;

        //        GlossaryDocument GlossaryDoc =
        //            mdpTemplate.GetPartsOfType<GlossaryDocumentPart>().FirstOrDefault().GlossaryDocument;

        //        var gDocPartBody = from x in GlossaryDoc.DocParts
        //                           where x.Descendants<DocPartProperties>().FirstOrDefault().DocPartName.Val == AutoTextName
        //                           select x;

        //        foreach (var docsubpart in gDocPartBody)
        //        {
        //            Drawing dr = docsubpart.Descendants<Drawing>().FirstOrDefault();
        //            if (dr != null)
        //            {
        //                relid = dr.Descendants<Blip>().FirstOrDefault().Embed.Value;
        //                Console.WriteLine("Relationship ID = {0}", relid);

        //                //var ImagePartandRel = from x in mdpTemplate.Parts
        //                //                      where x.RelationshipId == relid
        //                //                      select x;
        //            }
        //        }

        //        // this code means nothing because the relid, Relationship ID is for the Glossary Part 
        //        // and won't be the SAME in the actual template OR the new document
        //        Console.WriteLine("Uri = {0}",mdpTemplate.GetPartById(relid).Uri);
        //        Console.ReadLine();

        //        /* How do I retrieve an image part from the glossary document?
        //         * What does that look like in code and as xml?
        //         * How do I replace the current Image Part AND the run with Autotext?
        //         * Ideally it's:
        //         * 1. get DocPart from Glossary
        //         * 2. Insert it into new document
        //         * 3. Get relationship of Image saved with the DocPart
        //         * 4. Replace (or add) the new relationship 
        //         */

        //    }
        //}
    }
}
