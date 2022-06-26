using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace MetadataCleaner
{
    public class DocumentCleaner
    {
        private static readonly List<string> DOC_EXTENSIONS = new List<string> { ".docx", ".docm", ".dotx", ".dotm" };
        private static readonly List<string> SPREASHEET_EXTENSIONS = new List<string> { ".xlsx", ".xlsm", ".xltx", ".xltm", ".xlsb", ".xlam" };
        private static readonly List<string> PRESENTATION_EXTENSIONS = new List<string> { ".pptx", ".pptm", ".potx", ".potm", ".ppam", ".ppsx", ".ppsm", ".sldx", ".sldm", ".thmx" };

        public static void Clean(string pathToDocument)
        {
            string extension = Path.GetExtension(pathToDocument);

            if (DOC_EXTENSIONS.Contains(extension))
            {
                WordCleaner.Clean(pathToDocument);
            }
            else if (SPREASHEET_EXTENSIONS.Contains(extension))
            {
                ExcelCleaner.Clean(pathToDocument);
            }
            else if (PRESENTATION_EXTENSIONS.Contains(extension))
            {
                PowerPointCleaner.Clean(pathToDocument);
            }
            else
            {
                throw new Exception("The given file type (" + extension + ") is not supported.");
            }
        }// End Clean() method. 

    }
}
