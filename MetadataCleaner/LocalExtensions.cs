using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace MetadataCleaner
{
    public static class LocalExtensions
    {
        private static readonly XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        public static XDocument GetXDocument(this OpenXmlPart part)
        {
            XDocument xdoc = part.Annotation<XDocument>();
            if (xdoc != null)
                return xdoc;
            using (StreamReader sr = new StreamReader(part.GetStream()))
            using (System.Xml.XmlReader xr = XmlReader.Create(sr))
                xdoc = XDocument.Load(xr);
            part.AddAnnotation(xdoc);
            return xdoc;
        }
    }
}
