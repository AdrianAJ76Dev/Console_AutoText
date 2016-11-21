using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Wordprocessing;
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
        public static void DisplayInnerXML(string AutoTextName)
        {
            string relid;
            string templatefullname = @"C:\Users\ajones\Documents\Automation\Prototypes\Sole Source Letter\Templates\Sole Source Letter v4.dotx";
            using (WordprocessingDocument wrdTemplate = WordprocessingDocument.Open(templatefullname, false))
            {
                MainDocumentPart mdpTemplate = wrdTemplate.MainDocumentPart;
                Document doc = mdpTemplate.Document;

                GlossaryDocument GlossaryDoc =
                    mdpTemplate.GetPartsOfType<GlossaryDocumentPart>().FirstOrDefault().GlossaryDocument;

                var gDocPartBody = from x in GlossaryDoc.DocParts
                                   where x.Descendants<DocPartProperties>().FirstOrDefault().DocPartName.Val == AutoTextName
                                   select x;

                foreach (var docsubpart in gDocPartBody)
                {
                    Drawing dr = docsubpart.Descendants<Drawing>().FirstOrDefault();
                    if (dr != null)
                    {
                        Console.WriteLine("Drawing InnerXML = {0}", dr.InnerXml);
                        relid = dr.Descendants<Blip>().FirstOrDefault().Embed;
                        Console.WriteLine("Relationship ID = {0}", relid);

                        var ImagePartandRel = from x in mdpTemplate.Parts
                                              where x.RelationshipId == relid
                                              select x;
                    }
                }

                Console.ReadLine();

            }
        }
    }
}
