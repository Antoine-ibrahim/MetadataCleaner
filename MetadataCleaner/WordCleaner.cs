using System;
using System.IO;
using System.Xml;
using System.Linq;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Xml.Linq;
using OpenXmlPowerTools;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.ExtendedProperties;
using System.Configuration;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using System.IO.Compression;


namespace MetadataCleaner
{
    public class WordCleaner
    {
        #region Static Variables

        // Namespaces for various Word XML parts.
        private static readonly XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private static readonly XNamespace x = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
        private static readonly XNamespace cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
        private static readonly XNamespace dc = "http://purl.org/dc/elements/1.1/";
        private static readonly XNamespace dcterms = "http://purl.org/dc/terms/";

        // All possible file types using OpenXML.
        private static readonly List<string> DOC_EXTENSIONS = new List<string> { ".docx", ".docm", ".dotx", ".dotm" };
        private static readonly List<string> SPREASHEET_EXTENSIONS = new List<string> { ".xlsx", ".xlsm", ".xltx", ".xltm", ".xlsb", ".xlam" };
        private static readonly List<string> PRESENTATION_EXTENSIONS = new List<string> { ".pptx", ".pptm", ".potx", ".potm", ".ppam", ".ppsx", ".ppsm", ".sldx", ".sldm", ".thmx" };

        #endregion

        #region Public Methods

        //==== Clean ===================================================
        /// <summary>
        /// Cleans all metadata from the document.
        /// </summary>
        /// <param name="pathToDocument">The full path of the document that is being processed.</param>
        //=========================================================================
        public static void Clean(string pathToDocument)
        {
            string extension = Path.GetExtension(pathToDocument);

            if (DOC_EXTENSIONS.Contains(extension))
            {
                using (WordprocessingDocument wdDoc = getWordprocessingDocument(pathToDocument))
                {
                    try
                    {
                        // Get XDocument to perform direct XML edit operations with.
                        XDocument coreXDoc = wdDoc.CoreFilePropertiesPart.GetXDocument();

                        doRemoveSystemProperties(wdDoc, coreXDoc);
                        doRemoveBuiltInProperties(wdDoc, coreXDoc);
                        doRemoveTemplate(wdDoc);
                        doRemoveTracking(wdDoc);
                        doRemoveHiddenTextDoc(wdDoc);
                        doRemoveCommentsDoc(wdDoc);
                        doRemoveVariables(wdDoc);
                        doRemoveHyperlinks(wdDoc);
                        doRemoveIncludeFields(wdDoc);
                        doRemoveSmallFont(wdDoc);
                        doRemoveInvisibleText(wdDoc);
                        doRemoveHiddenObjects(wdDoc);
                        doRemovePersonalInfo(wdDoc, coreXDoc);
                        ////doRemoveControls(wdDoc);              // Only stubbed
                        // doRemoveCustomXML                      // Only stubbed
                        ////doRemoveGraphics(wdDoc);              // Only stubbed
                        doRemoveStatistics(wdDoc, coreXDoc);
                        doRemoveCustomProperties(wdDoc);
                        doRemoveMailMerge(wdDoc);
                        doReplaceEmbedsWithPictures(wdDoc);


                        // Save the document XML back to their document parts.
                        coreXDoc.Save(wdDoc.CoreFilePropertiesPart.GetStream(FileMode.Create, FileAccess.Write));
                    }
                    catch (Exception ex)
                    {
                        if (ex.Message.Contains("System.OutOfMemoryException"))
                        {
                            throw new FilePossiblyTooBiGException(ex.Message);
                        }
                        if (ex.GetType().Name == "MissingFilePropertiesException")
                        {
                            throw new MissingFilePropertiesException("Word");
                        }
                        else
                        {
                            throw new Exception(ex.Message);
                        }
                    }
                }
            }
            else if (SPREASHEET_EXTENSIONS.Contains(extension))
            {
                // TODO: Replace this line with spreadsheet handling code.
                throw new Exception("The given file type (" + extension + ") is not supported.");
            }
            else if (PRESENTATION_EXTENSIONS.Contains(extension))
            {
                // TODO: Replace this line with presentation handling code.
                throw new Exception("The given file type (" + extension + ") is not supported.");
            }
            else
            {
                throw new Exception("The given file type (" + extension + ") is not supported.");
            }
        }// End Clean() method. 

        #region Individual Public Methods



        //==== RemoveStatistics ===================================================
        /// <summary>
        /// Removes document statistics, such as time worked on the document and revision information.
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================
        public static void RemoveStatistics(string path)
        {
            using (WordprocessingDocument wdDoc = getWordprocessingDocument(path))
            {
                try
                {
                    XDocument coreXDoc = wdDoc.CoreFilePropertiesPart.GetXDocument();
                    doRemoveStatistics(wdDoc, coreXDoc);

                    coreXDoc.Save(wdDoc.CoreFilePropertiesPart.GetStream(FileMode.Create, FileAccess.Write));
                }
                catch (Exception ex)
                {
                    if (ex.Message.Contains("System.OutOfMemoryException"))
                    {
                        throw new FilePossiblyTooBiGException(ex.Message);
                    }
                    else
                    {
                        throw new Exception(ex.Message);
                    }
                }
            }
        }// End removeStatistics 


        //==== RemoveSystemProperties ===================================================
        /// <summary>
        /// This function removes all properties that are recorded by the system automatically, such as dates, authors, last modified etc...  
        /// </summary>
        /// <param name="path">This is the full path of the document that we are trying to remove Metadata From.</param>
        //=========================================================================

        public static void RemoveSystemProperties(string path)
        {
            using (WordprocessingDocument wdDoc = getWordprocessingDocument(path))
            {
                try
                {
                    XDocument coreXDoc = wdDoc.CoreFilePropertiesPart.GetXDocument();
                    doRemoveSystemProperties(wdDoc, coreXDoc);

                    coreXDoc.Save(wdDoc.CoreFilePropertiesPart.GetStream(FileMode.Create, FileAccess.Write));
                }
                catch (Exception ex)
                {
                    if (ex.Message.Contains("System.OutOfMemoryException"))
                    {
                        throw new FilePossiblyTooBiGException(ex.Message);
                    }
                    else
                    {
                        throw new Exception(ex.Message);
                    }
                }
            }
        } // End RemoveSystemProperties



        //==== RemoveBuiltInProperties ===================================================
        /// <summary>
        /// This function removes all built in properties located under the file tab.  
        /// </summary>
        /// <param name="path">This is the full path of the document that we are trying to remove Metadata From.</param>
        //=========================================================================

        public static void RemoveBuiltInProperties(string path)
        {
            using (WordprocessingDocument wdDoc = getWordprocessingDocument(path))
            {
                try
                {
                    XDocument coreXDoc = wdDoc.CoreFilePropertiesPart.GetXDocument();
                    doRemoveBuiltInProperties(wdDoc, coreXDoc);

                    coreXDoc.Save(wdDoc.CoreFilePropertiesPart.GetStream(FileMode.Create, FileAccess.Write));
                }
                catch (Exception ex)
                {
                    if (ex.Message.Contains("System.OutOfMemoryException"))
                    {
                        throw new FilePossiblyTooBiGException(ex.Message);
                    }
                    else
                    {
                        throw new Exception(ex.Message);
                    }
                }
            }
        } // End RemoveBuiltinProperties




        //==== RemoveTemplate ===================================================
        /// <summary>
        /// This function removes attached templates, all references to the attached template, and sets the word template back to the default Normal template.  
        /// </summary>
        /// <param name="path">This is the full path of the document that we are trying to remove Metadata From.</param>
        //=========================================================================
        public static void RemoveTemplate(string path)
        {
            using (WordprocessingDocument wdDoc = getWordprocessingDocument(path))
            {
                try
                {
                    doRemoveTemplate(wdDoc);
                }
                catch (Exception ex)
                {
                    if (ex.Message.Contains("System.OutOfMemoryException"))
                    {
                        throw new FilePossiblyTooBiGException(ex.Message);
                    }
                    else
                    {
                        throw new Exception(ex.Message);
                    }
                }
            }
        }// end RemoveTemplate



        //==== RemoveTracking ===================================================
        /// <summary>
        /// Approves all revisions, including deleted items, and edited items. If revision tracking is 
        /// on, it will be turned off 
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================
        public static void RemoveTracking(string path)
        {
            using (WordprocessingDocument wdDoc = getWordprocessingDocument(path))
            {
                try
                {
                    doRemoveTracking(wdDoc);
                }
                catch (Exception ex)
                {
                    if (ex.Message.Contains("System.OutOfMemoryException"))
                    {
                        throw new FilePossiblyTooBiGException(ex.Message);
                    }
                    else
                    {
                        throw new Exception(ex.Message);
                    }
                }
            }
        }// End RemoveTracking



        //==== RemoveHiddenText ===================================================
        /// <summary>
        /// This removes text that is invisible because it is set to hidden.
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================
        public static void RemoveHiddenText(string path)
        {
            using (WordprocessingDocument wdDoc = getWordprocessingDocument(path))
            {
                try
                {
                    doRemoveHiddenTextDoc(wdDoc);
                }
                catch (Exception ex)
                {
                    if (ex.Message.Contains("System.OutOfMemoryException"))
                    {
                        throw new FilePossiblyTooBiGException(ex.Message);
                    }
                    else
                    {
                        throw new Exception(ex.Message);
                    }
                }
            }
        }//End RemoveHiddenText



        //==== RemoveComments ===================================================
        /// <summary>
        /// Removes all comments from the document and removes all references to the authors of the comments.
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================
        public static void RemoveComments(string path)
        {
            using (WordprocessingDocument wdDoc = getWordprocessingDocument(path))
            {
                try
                {
                    doRemoveCommentsDoc(wdDoc);
                }
                catch (Exception ex)
                {
                    if (ex.Message.Contains("System.OutOfMemoryException"))
                    {
                        throw new FilePossiblyTooBiGException(ex.Message);
                    }
                    else
                    {
                        throw new Exception(ex.Message);
                    }
                }
            }
        }// End RemoveComments



        //public static void RemoveGraphics(string path)
        //{
        //    using (WordprocessingDocument wdDoc = getWordprocessingDocument(path))
        //    {
        //        try
        //        {
        //            doRemoveGraphics(wdDoc);
        //        }
        //        catch (Exception ex)
        //        {
        //            if (ex.Message.Contains("System.OutOfMemoryException"))
        //            {
        //                throw new FilePossiblyTooBiGException(ex.Message);
        //            }
        //            else
        //            {
        //                throw new Exception(ex.Message);
        //            }
        //        }
        //    }
        //}



        //==== RemoveVariables ===================================================
        /// <summary>
        /// Removes variables within the document, for example variables created with macros.
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================
        public static void RemoveVariables(string path)
        {
            using (WordprocessingDocument wdDoc = getWordprocessingDocument(path))
            {
                try
                {
                    doRemoveVariables(wdDoc);
                }
                catch (Exception ex)
                {
                    if (ex.Message.Contains("System.OutOfMemoryException"))
                    {
                        throw new FilePossiblyTooBiGException(ex.Message);
                    }
                    else
                    {
                        throw new Exception(ex.Message);
                    }
                }
            }
        }// End RemoveVariables



        //==== RemovePersonalInfo ===================================================
        /// <summary>
        /// Removes all personal info from the document and adds a setting that automatically removes personal information if the file is edited in the future.
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================
        public static void RemovePersonalInfo(string path)
        {
            using (WordprocessingDocument wdDoc = getWordprocessingDocument(path))
            {
                try
                {
                    XDocument coreXDoc = wdDoc.CoreFilePropertiesPart.GetXDocument();
                    doRemovePersonalInfo(wdDoc, coreXDoc);

                    coreXDoc.Save(wdDoc.CoreFilePropertiesPart.GetStream(FileMode.Create, FileAccess.Write));
                }
                catch (Exception ex)
                {
                    if (ex.Message.Contains("System.OutOfMemoryException"))
                    {
                        throw new FilePossiblyTooBiGException(ex.Message);
                    }
                    else
                    {
                        throw new Exception(ex.Message);
                    }
                }
            }
        }// End RemovePersonalInfo



        //==== RemoveHyperlinks ===================================================
        /// <summary>
        /// Remove any hyperlinks or on click links. This includes items such as a hyperlinked image.
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================
        public static void RemoveHyperlinks(string path)
        {
            using (WordprocessingDocument wdDoc = getWordprocessingDocument(path))
            {
                try
                {
                    doRemoveHyperlinks(wdDoc);
                }
                catch (Exception ex)
                {
                    if (ex.Message.Contains("System.OutOfMemoryException"))
                    {
                        throw new FilePossiblyTooBiGException(ex.Message);
                    }
                    else
                    {
                        throw new Exception(ex.Message);
                    }
                }
            }
        } // End RemoveHyperlinks


        //==== RemoveIncludeFields ===================================================
        /// <summary>
        /// Removes any references to include text fields, and replaces text with static text.
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================
        public static void RemoveIncludeFields(string path)
        {
            using (WordprocessingDocument wdDoc = getWordprocessingDocument(path))
            {
                try
                {
                    doRemoveIncludeFields(wdDoc);
                }
                catch (Exception ex)
                {
                    if (ex.Message.Contains("System.OutOfMemoryException"))
                    {
                        throw new FilePossiblyTooBiGException(ex.Message);
                    }
                    else
                    {
                        throw new Exception(ex.Message);
                    }
                }
            }
        } // End RemoveIncludeFields


        //==== RemoveSmallFont ===================================================
        /// <summary>
        /// Remove any text that is invisible or unnoticable because it has a font size that is 1 or smaller.
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================
        public static void RemoveSmallFont(string path)
        {
            using (WordprocessingDocument wdDoc = getWordprocessingDocument(path))
            {
                try
                {
                    doRemoveSmallFont(wdDoc);
                }
                catch (Exception ex)
                {
                    if (ex.Message.Contains("System.OutOfMemoryException"))
                    {
                        throw new FilePossiblyTooBiGException(ex.Message);
                    }
                    else
                    {
                        throw new Exception(ex.Message);
                    }
                }

            }
        }// End RemoveSmallFont



        //==== RemoveInvisibleText ===================================================
        /// <summary>
        /// This removes any text that has is the same color as the background. Ex red text that is highlighted red or has a background that is red
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================
        public static void RemoveInvisibleText(string path)
        {
            using (WordprocessingDocument wdDoc = getWordprocessingDocument(path))
            {
                try
                {
                    doRemoveInvisibleText(wdDoc);
                }
                catch (Exception ex)
                {
                    if (ex.Message.Contains("System.OutOfMemoryException"))
                    {
                        throw new FilePossiblyTooBiGException(ex.Message);
                    }
                    else
                    {
                        throw new Exception(ex.Message);
                    }
                }

            }
        } //end RemoveInvisibleText

        //public static void RemoveControls(string path)
        //{
        //    using (WordprocessingDocument wdDoc = getWordprocessingDocument(path))
        //    {
        //        try
        //        {
        //            doRemoveControls(wdDoc);
        //        }
        //        catch (Exception ex)
        //        {
        //            if (ex.Message.Contains("System.OutOfMemoryException"))
        //            {
        //                throw new FilePossiblyTooBiGException(ex.Message);
        //            }
        //            else
        //            {
        //                throw new Exception(ex.Message);
        //            }
        //        }
        //    }
        //}// end RemoveControls



        //==== RemoveHiddenObjects ===================================================
        /// <summary>
        /// This removes all hidden objects in the document, for example hidden pictures or text boxes.   
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================
        public static void RemoveHiddenObjects(string path)
        {
            using (WordprocessingDocument wdDoc = getWordprocessingDocument(path))
            {
                try
                {
                    doRemoveHiddenObjects(wdDoc);
                }
                catch (Exception ex)
                {
                    if (ex.Message.Contains("System.OutOfMemoryException"))
                    {
                        throw new FilePossiblyTooBiGException(ex.Message);
                    }
                    else
                    {
                        throw new Exception(ex.Message);
                    }
                }
            }
        }// End RemoveHiddenObjects



        //==== RemoveMailMerge ===================================================
        /// <summary>
        /// This removes all references to MailMerge data, and replaces it with text.   
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================
        public static void RemoveMailMerge(string path)
        {
            using (WordprocessingDocument wdDoc = getWordprocessingDocument(path))
            {
                try
                {
                    doRemoveMailMerge(wdDoc);
                }
                catch (Exception ex)
                {
                    if (ex.Message.Contains("System.OutOfMemoryException"))
                    {
                        throw new FilePossiblyTooBiGException(ex.Message);
                    }
                    else
                    {
                        throw new Exception(ex.Message);
                    }
                }
            }
        }// End RemoveMailMerge


        //==== RemoveCustomProperties ===================================================
        /// <summary>
        /// Removes all custom properties that were added to the document by users.   
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================
        public static void RemoveCustomProperties(string path)
        {
            using (WordprocessingDocument wdDoc = getWordprocessingDocument(path))
            {
                try
                {
                    doRemoveCustomProperties(wdDoc);
                }
                catch (Exception ex)
                {
                    if (ex.Message.Contains("System.OutOfMemoryException"))
                    {
                        throw new FilePossiblyTooBiGException(ex.Message);
                    }
                    else
                    {
                        throw new Exception(ex.Message);
                    }
                }
            }

        } // End RemoveCustomProperties



        //==== ReplaceEmbedsWithPictures ===================================================
        /// <summary>
        /// Replaces all embedded objects with a picture and removes all references to any embedded documents.  
        /// </summary>
        /// <param name="Path">The full path of the document that is being processed.</param>
        //=========================================================================
        public static void ReplaceEmbedsWithPictures(string path)
        {
            using (WordprocessingDocument wdDoc = getWordprocessingDocument(path))
            {
                try
                {
                    doReplaceEmbedsWithPictures(wdDoc);
                }
                catch (Exception ex)
                {
                    if (ex.Message.Contains("System.OutOfMemoryException"))
                    {
                        throw new FilePossiblyTooBiGException(ex.Message);
                    }
                    else
                    {
                        throw new Exception(ex.Message);
                    }
                }
            }
        } //End ReplaceEmbedsWithPicture


        //==== UnCheckCheckboxes ===================================================
        /// <summary>
        /// This removes text that is invisible because it is set to hidden.
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================
        public static void UnCheckCheckboxes(string path)
        {
            using (WordprocessingDocument wdDoc = getWordprocessingDocument(path))
            {
                try
                {
                    doUncheckCheckboxs(wdDoc);
                }
                catch (Exception ex)
                {
                    if (ex.Message.Contains("System.OutOfMemoryException"))
                    {
                        throw new FilePossiblyTooBiGException(ex.Message);
                    }
                    else
                    {
                        throw new Exception(ex.Message);
                    }
                }
            }
        }//End RemoveHiddenText


        #endregion

        #endregion

        #region Private Methods


        //==== doRemoveTemplate ===================================================
        /// <summary>
        /// This function removes attached templates, all references to the attached template, and sets the word template back to the default Normal template.  
        /// </summary>
        /// <param name="path">This is the full path of the document that we are trying to remove Metadata From.</param>
        //=========================================================================

        private static void doRemoveTemplate(WordprocessingDocument wDoc)
        {
            // Removes the template.

            MainDocumentPart mainPart = wDoc.MainDocumentPart;
            // get the template from the document properties, and get the settings attachment connection from the settings  
            OpenXmlElement attachedTemplate = mainPart.DocumentSettingsPart.Settings
                .Where(a => a.LocalName == "attachedTemplate").FirstOrDefault();
            OpenXmlElement TemplateProperty = wDoc.ExtendedFilePropertiesPart.Properties
                .Where(a => a.LocalName == "Template").FirstOrDefault();

            if (attachedTemplate != null)
            {
                // get the id of the attachedTemplate 
                string id = attachedTemplate.GetAttributes()
                    .Where(a => a.LocalName.ToString() == "id").FirstOrDefault().Value.ToString();

                // Delete both the attached Template and the referenced relationship of the attached template from the document settings 
                ReferenceRelationship relationship = mainPart.DocumentSettingsPart.GetReferenceRelationship(id);
                mainPart.DocumentSettingsPart.DeleteReferenceRelationship(relationship);
                attachedTemplate.Remove();
            }

            // replace the template reference to specify the Normal template
            Template template1 = new Template();
            template1.Text = "Normal.dotm";
            if (TemplateProperty != null)
            {
                TemplateProperty.InsertAfterSelf(template1);
                TemplateProperty.Remove();
            }



        }// End doRemoveTemplate



        //==== doRemoveTracking ===================================================
        /// <summary>
        /// Appreoves all revisions, including deleted items, and edited items. If revision tracking is 
        /// on, it will be turned off 
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================
        private static void doRemoveTracking(WordprocessingDocument wdDoc)
        {
            MainDocumentPart mainPart = wdDoc.MainDocumentPart;
            //TODO; update the if statements 
            // Removes change tracking by accepting pending changes.
            if (!RevisionAccepter.HasTrackedRevisions(wdDoc))
            {
                //Console.WriteLine("No change tracking found.");
            }
            else
            {
                RevisionAccepter.AcceptRevisions(wdDoc);
                List<OpenXmlElement> EmptyTables = mainPart.Document.Body.Descendants()
                    .Where(a => a.LocalName == "tbl"
                        && a.Descendants().Where(b => b.LocalName == "tr").Count() < 1).ToList();
                if (EmptyTables != null)
                {
                    foreach (OpenXmlElement EmptyTable in EmptyTables)
                    {
                        EmptyTable.Remove();
                    }
                }
            }
            // check if revision tracking is turned on , if so remove it 
            OpenXmlElement revisionTrackingElement = mainPart.DocumentSettingsPart.Settings
                .Where(a => a.LocalName == "trackRevisions").FirstOrDefault();
            if (revisionTrackingElement != null)
            {
                revisionTrackingElement.Remove();
            }
        } //End doRemoveTracking



        //==== doRemoveHiddenText ===================================================
        /// <summary>
        /// This removes text that is invisible because it is set to hidden.
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================
        private static void doRemoveHiddenTextDoc(WordprocessingDocument wDoc)
        {
            //TODO: remove if 
            MainDocumentPart mainPart = wDoc.MainDocumentPart;
            List<OpenXmlElement> hiddenNodes = mainPart.Document.Body.Descendants<OpenXmlElement>()
                .Where(a => a.LocalName == "vanish").ToList();

            if (hiddenNodes.Count == 0)
            {
                //Console.WriteLine("No hidden text found.");
                return;
            }
            else
            {
                foreach (OpenXmlElement hiddenNode in hiddenNodes)
                {
                    OpenXmlElement topNode = hiddenNode.Parent.Parent;
                    topNode.Remove();

                }
            }

        } // End doRemoveHiddenText




        //==== doRemoveComments ===================================================
        /// <summary>
        /// Removes all comments from the document and removes all references to the authors of the comments.
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================
        private static void doRemoveCommentsDoc(WordprocessingDocument wdDoc)
        {
            //TODO
            MainDocumentPart mainPart = wdDoc.MainDocumentPart;
            // Set commentPart to the document WordprocessingCommentsPart, if it exists.
            WordprocessingCommentsPart commentPart = mainPart.WordprocessingCommentsPart;
            WordprocessingCommentsExPart commentExPart = mainPart.WordprocessingCommentsExPart;
            if (commentPart == null && commentExPart == null)
            {
                //Console.WriteLine("No comments found.");
            }
            else
            {
                List<OpenXmlElement> commentRefs = mainPart.Document.Body.Descendants()
                    .Where(a => a.LocalName == "commentReference").ToList();
                List<OpenXmlElement> CommentMarkers = mainPart.Document.Body.Descendants()
                    .Where(a => (a.LocalName == "commentRangeEnd" || a.LocalName == "commentRangeStart")).ToList();

                foreach (OpenXmlElement Marker in CommentMarkers)
                {
                    Marker.Remove();
                }
                foreach (OpenXmlElement refrence in commentRefs)
                {
                    refrence.Parent.Remove();
                }

                if (commentPart != null)
                {
                    mainPart.DeletePart(commentPart);
                }
                if (commentExPart != null)
                {
                    mainPart.DeletePart(commentExPart);
                }

                // remove the authors for the comments 
                WordprocessingPeoplePart peoplePart = mainPart.WordprocessingPeoplePart;

                // CHeck if there are authors for the comments and remove those also .
                if (peoplePart == null)
                {
                    //Console.WriteLine("No authors found.");
                }
                else
                {
                    // Remove the people part.
                    wdDoc.MainDocumentPart.DeletePart(peoplePart);
                }

            }
        } // End doRemoveComments

        private static void doRemoveGraphics(WordprocessingDocument wDoc)
        {
            // Removes graphics in document.
            // Code goes here.
        }//End doRemoveGraphics



        //==== doRemoveVariables ===================================================
        /// <summary>
        /// removes variables within the document, for example variables created with macros.
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================
        private static void doRemoveVariables(WordprocessingDocument wDoc)
        {
            // Removes document variables (VBA). These are variables that were inserted into the document using the developer tab to create macros 

            MainDocumentPart mainPart = wDoc.MainDocumentPart;

            OpenXmlElement DocVariables = mainPart.DocumentSettingsPart.Settings
                .Where(a => a.LocalName == "docVars").FirstOrDefault();

            if (DocVariables != null)
            {
                DocVariables.Remove();
            }


        }// End doRemoveVariables



        //==== doRemoveHyperlinks ===================================================
        /// <summary>
        /// Remove any hyperlinks or on click links. This includes items such as a hyperlinked image.
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================
        private static void doRemoveHyperlinks(WordprocessingDocument wDoc)
        {
            // Removes hyperlinks from document.

            MainDocumentPart mainPart = wDoc.MainDocumentPart;

            // get and remove all hyperlink elements 
            List<OpenXmlElement> ALlHyperlinks = mainPart.Document.Body.Descendants<OpenXmlElement>()
                .Where(a => a.LocalName == "hyperlink").ToList();
            foreach (OpenXmlElement hyperlink in ALlHyperlinks)
            {
                string id = hyperlink.GetAttributes()
                    .Where(a => a.LocalName == "id").FirstOrDefault().Value;
                if (id != null)
                {
                    ReferenceRelationship hlr = mainPart.GetReferenceRelationship(id);

                    Run run = new Run();
                    Text text = new Text();
                    text.Text = hyperlink.InnerText;
                    run.Append(text);
                    hyperlink.InsertAfterSelf(run);

                    // remove the hyperlink, and its relationship 
                    hyperlink.Remove();
                    mainPart.DeleteReferenceRelationship(hlr);
                }
            }

            FootnotesPart footNotespart= mainPart.FootnotesPart;
            if (footNotespart != null)
            {


                List<OpenXmlElement> footNotesHyperlinks = mainPart.FootnotesPart.Footnotes.Descendants<OpenXmlElement>().Where(a => a.LocalName == "hyperlink").ToList();
                foreach (OpenXmlElement link in footNotesHyperlinks)
                {
                    string id = link.GetAttributes()
                        .Where(a => a.LocalName == "id").FirstOrDefault().Value;
                    if (id != null)
                    {
                        ReferenceRelationship rel = mainPart.FootnotesPart.GetReferenceRelationship(id);

                        Run run = new Run();
                        Text text = new Text();
                        text.Text = link.InnerText;
                        run.Append(text);
                        link.InsertAfterSelf(run);

                        link.Remove();
                        mainPart.FootnotesPart.DeleteReferenceRelationship(rel);
                    }

                }

            }


            // Sometimes hyperlinks are added to the xml as fieldcode elements instead of hyperlink elements, making them undetectable using the above method 
            // gets all the hyperlinks disguised as fieldcode elements in the xml
            List<OpenXmlElement> LinkedFieldCodes = mainPart.Document.Body.Descendants<OpenXmlElement>()
                    .Where(a => a.GetType().Name.ToString() == "FieldCode"
                    && a.InnerText.ToUpper().Contains("HYPERLINK")).ToList();

            foreach (OpenXmlElement field in LinkedFieldCodes)
            {
                field.Remove();
            }

            // gets all the remaining style elements indicate the text is a hyperlink. The remaining elements will include elements that were
            // missed due to disguised hyperlinks as fieldcode elements  
            IEnumerable<OpenXmlElement> RunStyleLinks = mainPart.Document.Body.Descendants()
                .Where(a => a.LocalName == "rStyle"
                && a.GetAttributes().Where(b => b.LocalName.ToString() == "val").FirstOrDefault().Value == "Hyperlink");

            foreach (OpenXmlElement RunStyleElement in RunStyleLinks)
            {
                RunStyleElement.Remove();
            }

            // remove HLinkClick items, which are hyperlinks added to other elements such as pictures, objects, etc...
            List<OpenXmlElement> AllHLinkClicks = mainPart.Document.Body.Descendants<OpenXmlElement>()
                .Where(a => a.LocalName == "hlinkClick").ToList();
            foreach (OpenXmlElement Hlink in AllHLinkClicks)
            {
                bool ReferenceAlredyDeleted = false;
                string id = Hlink.GetAttributes()
                    .Where(a => a.LocalName == "id").FirstOrDefault().Value;
                ReferenceRelationship hlr = null;
                try
                {
                    hlr = mainPart.GetReferenceRelationship(id);
                }
                catch
                {
                    ReferenceAlredyDeleted = true;
                }

                if (ReferenceAlredyDeleted == false)
                {
                    Hlink.Remove();
                    mainPart.DeleteReferenceRelationship(hlr);
                }
                else
                {
                    Hlink.Remove();
                }
            }


        }// End doRemoveHyperlinks



        //==== doRemovePersonalInfo ===================================================
        /// <summary>
        /// Removes all personal info from the document and adds a setting that automatically removes personal information if the file is edited in the future.
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================

        private static void doRemovePersonalInfo(WordprocessingDocument wDoc, XDocument coreXDoc)
        {

            MainDocumentPart mainPart = wDoc.MainDocumentPart;
            ExtendedFilePropertiesPart ExtendedProperties = wDoc.ExtendedFilePropertiesPart;

            // Removes Author and last modified by.
            XName creator = dc + "creator";
            XName lastModified = cp + "lastModifiedBy";

            coreXDoc.Descendants()
                .Where(x => x.Name == lastModified || x.Name == creator)
                .Remove();

            // Removes the Manager property.
            OpenXmlElement ManagerSetting = ExtendedProperties.Properties
                .Where(a => a.LocalName == "Manager").FirstOrDefault();

            if (ManagerSetting != null)
            {
                Ap.Manager Manager = new Ap.Manager();
                ManagerSetting.InsertAfterSelf(Manager);
                ManagerSetting.Remove();
            }

            // Edits the author of all the comments to default "Author".
            WordprocessingCommentsPart commentPart = wDoc.MainDocumentPart.WordprocessingCommentsPart;


            if (commentPart != null)
            {
                foreach (Comment coment in commentPart.Comments)
                {
                    coment.Author.Value = "Author";
                    coment.Initials.Value = "A";
                }
            }
            // Removes the people part, which houses information about all people who have added comments. 
            WordprocessingPeoplePart peoplePart = mainPart.WordprocessingPeoplePart;

            // CHeck if there are authors for the comments and remove those also .
            if (peoplePart != null)
            {
                // Remove the people part.
                wDoc.MainDocumentPart.DeletePart(peoplePart);
            }


            // Adds the RemovePersonalInformation trigger which automatically removes personal information from future revisions. 
            OpenXmlElement RemovePersonalInfoTag = mainPart.DocumentSettingsPart.Settings
                .Where(a => a.LocalName == "removePersonalInformation").FirstOrDefault();

            if (RemovePersonalInfoTag == null)
            {
                Settings WordSettings = mainPart.DocumentSettingsPart.Settings;
                RemovePersonalInformation RemovePersonalInfo = new RemovePersonalInformation();
                WordSettings.Append(RemovePersonalInfo);
            }
        } // End doRemovePersonalInfo



        //====doRemoveIncludeFields ===================================================
        /// <summary>
        /// removes any references to include text fields, and replaces text with static text.
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================
        private static void doRemoveIncludeFields(WordprocessingDocument wDoc)
        {
            MainDocumentPart mainPart = wDoc.MainDocumentPart;

            List<OpenXmlElement> IncludeTextParagraphs = mainPart.Document.Body.Descendants()
                .Where(a => a.GetType().Name.ToString() == "Paragraph"
                && (a.InnerText.Contains("INCLUDETEXT")
                || a.InnerText.Contains("INCLUDEPICTURE"))).ToList();

            foreach (OpenXmlElement paragraph in IncludeTextParagraphs)
            {
                // will need to remove all these
                List<OpenXmlElement> fldchars = paragraph.Descendants()
                    .Where(a => a.GetType().Name.ToString() == "Run"
                    && a.Descendants().Where(b => b.LocalName == "fldChar"
                    && b.GetAttributes().Where(c => c.LocalName.ToString() == "fldCharType"
                    && (c.Value == "begin" || c.Value == "separate" || c.Value == "end")).FirstOrDefault().Value != null).Count() > 0).ToList();
                foreach (OpenXmlElement fldchar in fldchars)
                {
                    fldchar.Remove();
                }

                // was originally just deleting specific instrtext fields, now it is deleting all of them 
                List<OpenXmlElement> instrTexts = paragraph.Descendants()
                    .Where(a => a.GetType().Name.ToString() == "Run"
                    && a.Descendants().Where(b => b.LocalName == "instrText").Count() > 0).ToList();
                //  && (b.InnerText.Trim() =="" || b.InnerText.Contains("INCLUDETEXT") || b.InnerText.Contains("INCLUDEPICTURE"))).Count() >0).ToList();
                foreach (OpenXmlElement instrText in instrTexts)
                {
                    instrText.Remove();
                }


                List<OpenXmlElement> imageDatas = paragraph.Descendants()
                    .Where(a => a.LocalName == "imagedata"
                    && a.GetAttributes().Where(b => b.LocalName == "href").FirstOrDefault().Value != null).ToList();

                foreach (OpenXmlElement imageData in imageDatas)
                {
                    OpenXmlAttribute Href = imageData.GetAttributes()
                        .Where(b => b.LocalName == "href").FirstOrDefault();
                    string rId = imageData.GetAttributes()
                        .Where(b => b.LocalName == "id").FirstOrDefault().Value;

                    string hrefid = Href.Value;
                    string uri = Href.NamespaceUri;

                    ReferenceRelationship HrefRelationship = mainPart.GetReferenceRelationship(hrefid);
                    mainPart.DeleteReferenceRelationship(HrefRelationship);

                    DocumentFormat.OpenXml.Vml.ImageData imageData1 = new DocumentFormat.OpenXml.Vml.ImageData()
                    {
                        RelationshipId = rId
                    };

                    imageData.InsertAfterSelf(imageData1);
                    imageData.Remove();
                }

            }
        } // End doRemoveIncludeFields



        //==== doRemoveSmallFont ===================================================
        /// <summary>
        /// Remove any text that is invisible or unnoticable because it has a font size that is 1 or smaller.
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================
        private static void doRemoveSmallFont(WordprocessingDocument wDoc)
        {

            // TODO: add comments explaining pt sizes , catch elements remove exception if not handeling anything, update code to check for breaking element instead of catch 
            // Removes text with small font (1pt or smaller).
            MainDocumentPart mainPart = wDoc.MainDocumentPart;
            bool textRemoved = false;

            List<OpenXmlElement> runElements = mainPart.Document.Body.Descendants()
                .Where(a => a.GetType().Name == "Run"
                && a.Descendants().Where(b => b.LocalName == "rPr").FirstOrDefault() != null).ToList();


            foreach (OpenXmlElement r in runElements)
            {
                try
                {
                    OpenXmlElement runParent = r.Ancestors()
                        .Where(a => a.GetType().Name.ToString() == "Paragraph").FirstOrDefault();

                    string sizeElement = r.Descendants()
                        .Where(b => b.LocalName == "rPr").FirstOrDefault()
                        .Descendants().Where(a => a.LocalName == "sz").FirstOrDefault()
                        .GetAttributes().Where(a => a.LocalName.ToString() == "val").FirstOrDefault().Value;

                    if (Convert.ToInt32(sizeElement) <= 2)
                    {
                        // text tat is size 1 or less 
                        r.Remove();
                        textRemoved = true;
                    }

                    if (textRemoved = true && runParent.Descendants()
                        .Where(a => a.GetType().Name.ToString() == "Run").Count() < 1)
                    {
                        runParent.Remove();
                    }
                }
                catch (Exception)
                {
                    // steps in here if the text does not have a size value attribute, we do not need to do anything to handle this 
                    // as it is not an error, it is just 
                }
            }

            if (!textRemoved)
            {
                // Console.WriteLine("No small text found.");
                return;
            }
        } // End doRemoveSmallFont



        //==== doRemoveInvisibleText ===================================================
        /// <summary>
        /// This removes any text that has is the same color as the background. Ex red text that is highlighted red or has a background that is red
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================
        private static void doRemoveInvisibleText(WordprocessingDocument wDoc)
        {

            // TODO: check if payne has leverage for color differences
            // Removes font matching background color.
            MainDocumentPart mainPart = wDoc.MainDocumentPart;
            // Removes font matching background color.
            bool textRemoved = false;
            Dictionary<string, string> HighlightHex = new Dictionary<string, string>(); ;

            // highlights are not tested for because they dont have a hex representation, only the text name. Creating a dictionary
            // to also remove hidden text due to highlights

            HighlightHex.Add("yellow", "FFFF00");
            HighlightHex.Add("green", "00FF00");
            HighlightHex.Add("cyan", "00FFFF");
            HighlightHex.Add("magenta", "FF00FF");
            HighlightHex.Add("blue", "0000FF");
            HighlightHex.Add("red", "FF0000");
            HighlightHex.Add("darkBlue", "00008B");
            HighlightHex.Add("darkCyan", "008B8B");
            HighlightHex.Add("darkMagenta", "800080");
            HighlightHex.Add("darkRed", "8B0000");
            HighlightHex.Add("darkYellow", "808000");
            HighlightHex.Add("Black", "000000");
            HighlightHex.Add("darkGray", "A9A9A9B");
            HighlightHex.Add("lightGray", "D3D3D3");

            // get all paragraphs 

            // get and remove all hyperlink elements 
            List<OpenXmlElement> paragraphElements = mainPart.Document.Body.Descendants()
                .Where(a => a.GetType().Name.ToString() == "Paragraph").ToList();

            foreach (OpenXmlElement p in paragraphElements)
            {
                // get all the run elements that specify a text color 
                List<OpenXmlElement> runWithColor = p.Descendants()
                    .Where(a => a.GetType().Name.ToString() == "Run"
                    && a.Descendants().Where(b => b.LocalName == "color"
                    && b.GetAttributes().Where(c => c.LocalName.ToString() == "val").FirstOrDefault().Value != null).Count() > 0).ToList();

                bool removedText = false;

                foreach (OpenXmlElement runText in runWithColor)
                {
                    string backgroundColor = "";
                    string highlightColor = "";
                    string textcolor = runText.Descendants()
                        .Where(b => b.LocalName == "color").FirstOrDefault().GetAttributes()
                        .Where(c => c.LocalName.ToString() == "val").FirstOrDefault().Value;

                    OpenXmlElement HighlightElement = runText.Descendants()
                        .Where(b => b.LocalName == "highlight").FirstOrDefault();//.Attribute(val).Value;

                    OpenXmlElement runBgElement = runText.Descendants()
                        .Where(b => b.LocalName == "shd").FirstOrDefault();//.Attribute(val).Value;

                    OpenXmlElement parentBgElement = null;
                    if (runBgElement != null)
                    {
                        parentBgElement = runText.Parent.Descendants()
                           .Where(b => b.LocalName == "pPr").FirstOrDefault()
                           .Descendants().Where(b => b.LocalName == "shd").FirstOrDefault();
                    }

                    if (runBgElement != null || parentBgElement != null)
                    {
                        if (runBgElement != null)
                        {
                            backgroundColor = runBgElement.GetAttributes()
                                .Where(c => c.LocalName.ToString() == "fill").FirstOrDefault().Value;

                        }
                        else
                        {
                            backgroundColor = parentBgElement.GetAttributes()
                                .Where(c => c.LocalName.ToString() == "fill").FirstOrDefault().Value;
                        }
                        // Console.WriteLine("the text is " + textcolor + " and has a backgroundColor of " + backgroundColor);

                    }

                    // check if the text is highlighted get color
                    if (HighlightElement != null)
                    {
                        highlightColor = HighlightHex[HighlightElement.GetAttributes()
                            .Where(c => c.LocalName.ToString() == "val").FirstOrDefault().Value];

                        // Console.WriteLine("the text is " + textcolor + " and is highlighted in " + highlightColor);
                    }

                    if (textcolor == backgroundColor || textcolor == highlightColor)
                    {
                        // OpenXmlElement prevSibling = runText.PreviousSibling().FirstOrDefault();
                        runText.Remove();
                        removedText = true;
                        textRemoved = true;

                    }

                }


                if (removedText == true && p.Descendants()
                    .Where(a => a.GetType().Name.ToString() == "Run").ToList().Count < 1)
                {
                    p.Remove();
                }

            }

            if (!textRemoved)
            {
                // Console.WriteLine("No invisible text found.");
                return;
            }
        } // ENd doRemoveInvisibleText

        private static void doRemoveControls(WordprocessingDocument wDoc)
        {
            // Removes content controls.
            // Code goes here.
        }// End doRemoveControls



        //==== doRemoveHiddenObjects ===================================================
        /// <summary>
        /// This removes all hidden objects in the document, for example hidden pictures or text boxes.   
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================
        private static void doRemoveHiddenObjects(WordprocessingDocument wDoc)
        {
            // Removes hidden objects.                
            MainDocumentPart mainPart = wDoc.MainDocumentPart;

            //objects such as text boxes and shapes that can be interpreted differently by other applications are stored in the xml as alternate text
            // here we pull all the alternate text items that have a their "hidden" property set to "1"
            IEnumerable<OpenXmlElement> AlternateContents = mainPart.Document.Body.Descendants().Where(a => a.LocalName == "AlternateContent"
                && a.Descendants().Where(b => b.LocalName.ToString() == "docPr"
                && b.GetAttributes().Where(c => c.LocalName.ToString() == "hidden").FirstOrDefault().Value == "1").Count() > 0);

            foreach (OpenXmlElement content in AlternateContents)
            {
                content.Remove();

            }
        } // End doRemoveHiddenObjects



        //==== doRemoveCustomProperties ===================================================
        /// <summary>
        /// Removes all custom properties that were added to the document by users.   
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================
        private static void doRemoveCustomProperties(WordprocessingDocument wDoc)
        {
            // removes custom Properties
            CustomFilePropertiesPart CustomProperties = wDoc.CustomFilePropertiesPart;
            wDoc.DeletePart(CustomProperties);

        }// End doRemoveCustomProperties



        //==== doRemoveMailMerge ===================================================
        /// <summary>
        /// This removes all references to MailMerge data, and replaces it with text.   
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================
        private static void doRemoveMailMerge(WordprocessingDocument wDoc)
        {
            MainDocumentPart mainPart = wDoc.MainDocumentPart;
            // remove recipientdata.xml
            MailMergeRecipientDataPart merger = mainPart.DocumentSettingsPart.MailMergeRecipientDataPart;

            if (merger != null)
            {
                string mergerid = mainPart.DocumentSettingsPart.GetIdOfPart(merger);
                mainPart.DocumentSettingsPart.DeletePart(mergerid);

            }
            List<ExternalRelationship> external = mainPart.DocumentSettingsPart.ExternalRelationships
                .Where(a => a.RelationshipType.Contains("mailMergeSource")).ToList();
            foreach (ExternalRelationship rel in external)
            {
                mainPart.DocumentSettingsPart.DeleteExternalRelationship(rel);
            }
            // remove mailmerge info from word/rels/settings.xml.rels


            // remove mailmerge from settings.xml
            OpenXmlElement MailMergeData = mainPart.DocumentSettingsPart.Settings
                .Where(a => a.LocalName == "mailMerge").FirstOrDefault();
            if (MailMergeData != null)
            {
                MailMergeData.Remove();
            }

            // clean up document
            List<OpenXmlElement> MergeFields = mainPart.Document.Body.Descendants()
                .Where(a => a.LocalName == "fldSimple" && a.GetAttributes()
                .Where(b => b.LocalName == "instr").FirstOrDefault().Value.Contains("MERGEFIELD")).ToList();
            foreach (OpenXmlElement Field in MergeFields)
            {
                foreach (OpenXmlElement fieldChild in Field.ChildElements)
                {
                    OpenXmlElement ChildCopy = fieldChild.CloneNode(true);
                    Field.InsertBeforeSelf(ChildCopy);
                }

                Field.Remove();
            }

        } // End doRemoveMailMerge



        //==== doReplaceEmbedsWithPictures ===================================================
        /// <summary>
        /// Replaces all embedded objects with a picture and removes all references to any embedded documents.  
        /// </summary>
        /// <param name="Path">The full path of the document that is being processed.</param>
        //=========================================================================
        private static void doReplaceEmbedsWithPictures(WordprocessingDocument wDoc)
        {
            // replaces embedded objects that link to other documents into images
            MainDocumentPart mainPart = wDoc.MainDocumentPart;

            // get all the embeded object from the document
            List<OpenXmlElement> allEmbeds = mainPart.Document.Body.Descendants<OpenXmlElement>()
                .Where(a => a.GetType().Name.ToString() == "OleObject").ToList();

            // for every single embed, ...
            foreach (OpenXmlElement embed in allEmbeds)
            {
                // get the run parent element and run ancestor 
                OpenXmlElement EmbedRunParent = embed.Ancestors()
                    .Where(a => a.GetType().Name == "Run").FirstOrDefault();
                OpenXmlElement ObjectElemet = embed.Parent;
                string Doc_ID = "";

                Doc_ID = embed.GetAttributes()
                    .Where(a => a.LocalName == "id").FirstOrDefault().Value;

                // create a new picture element and append to it all the appropriate child elements from the embedded object
                Picture picture1 = new Picture();

                foreach (OpenXmlElement element in ObjectElemet)
                {
                    if (element.GetType().Name.ToString() != "OleObject")
                    {
                        OpenXmlElement elementCopy = element.CloneNode(true);
                        picture1.Append(elementCopy);
                    }
                }

                // replace the Embed with the created picture, and delete the reference to the Embedded document. 
                ObjectElemet.InsertAfterSelf(picture1);
                ObjectElemet.Remove();
                mainPart.DeletePart(Doc_ID);

            }

        } // End doReplaceEmbedsWithPictures




        //==== doRemoveStatistics ===================================================
        /// <summary>
        /// Removes document statistics, such as time worked on the document and revision information.
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================
        private static void doRemoveStatistics(WordprocessingDocument wDoc, XDocument coreXDoc)
        {
            MainDocumentPart mainPart = wDoc.MainDocumentPart;
            ExtendedFilePropertiesPart docProperties = wDoc.ExtendedFilePropertiesPart;

            OpenXmlElement totalTimes = docProperties.Properties
                .Where(a => a.LocalName == "TotalTime").FirstOrDefault();

            if (totalTimes != null)
            {
                Ap.TotalTime newTotalTimes = new Ap.TotalTime();
                newTotalTimes.Text = "0";

                //   IEnumerable<OpenXmlElement> CorePropertiesToDelete;
                totalTimes.InsertAfterSelf(newTotalTimes);
                totalTimes.Remove();

            }


            XName lastModified = cp + "lastModifiedBy";
            coreXDoc.Descendants()
                .Where(x => x.Name == lastModified)
                .Remove();

            XElement revision = coreXDoc.Elements(cp + "coreProperties").Elements(cp + "revision").FirstOrDefault();
            if (revision != null)
            {
                revision.Value = "1";
            }

        } // End doRemoveStatistics


        //==== doRemoveSystemProperties ===================================================
        /// <summary>
        /// This function removes all properties that are recorded by the system automatically, such as dates, authors, last modified etc...  
        /// </summary>
        /// <param name="path">This is the full path of the document that we are trying to remove Metadata From.</param>
        //=========================================================================

        private static void doRemoveSystemProperties(WordprocessingDocument wDoc, XDocument coreXDoc)
        {
            if (coreXDoc.Descendants().Count() <=2 || wDoc.ExtendedFilePropertiesPart.Properties.Count()<=1)
            {
                throw new MissingFilePropertiesException("Word");
            }

            MainDocumentPart mainPart = wDoc.MainDocumentPart;
            ExtendedFilePropertiesPart ex = wDoc.ExtendedFilePropertiesPart;
            CoreFilePropertiesPart cp = wDoc.CoreFilePropertiesPart;

            // change the appversion to 12.0000
            OpenXmlElement AppVersion = ex.Properties
                .Where(a => a.LocalName == "AppVersion").FirstOrDefault();
            Ap.ApplicationVersion applicationVersion = new Ap.ApplicationVersion();
            applicationVersion.Text = "12.0000";

            if (AppVersion != null)
            {
                AppVersion.InsertAfterSelf(applicationVersion);
                AppVersion.Remove();
            }
            else {
                ex.Properties.ApplicationVersion= applicationVersion;
            }

            // edit the created and modified date to the date that the cleaner was run.
            XName created = dcterms + "created";
            XName modified = dcterms + "modified";

            coreXDoc.Descendants()
                .Where(x => x.Name == created || x.Name == modified)
                .ToList()
                .ForEach(e => e.Value = DateTime.Now.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'"));

        } // End doRemoveSystemProperties




        //==== doRemoveBuiltInProperties ===================================================
        /// <summary>
        /// This function removes all built in properties located under the file tab.  
        /// </summary>
        /// <param name="path">This is the full path of the document that we are trying to remove Metadata From.</param>
        //=========================================================================
        private static void doRemoveBuiltInProperties(WordprocessingDocument wDoc, XDocument coreXDoc)
        {

            if (coreXDoc.Descendants().Count() <= 2 || wDoc.ExtendedFilePropertiesPart.Properties.Count() <= 1)
            {
                throw new MissingFilePropertiesException("Word");
            }

            MainDocumentPart mainPart = wDoc.MainDocumentPart;
            ExtendedFilePropertiesPart ex = wDoc.ExtendedFilePropertiesPart;

            OpenXmlElement company = ex.Properties
                .Where(a => a.LocalName == "Company").FirstOrDefault();

            OpenXmlElement hLinkBase = ex.Properties
                .Where(a => a.LocalName == "HyperlinkBase").FirstOrDefault();
            //TODO: edit variables to not have numbers
            Ap.Company company1 = new Ap.Company();
            if (company != null)
            {
                company.InsertAfterSelf(company1);
                company.Remove();
            }


            Ap.HyperlinkBase hyperlinkBase1 = new Ap.HyperlinkBase();
            if (hLinkBase != null)
            {
                hLinkBase.InsertAfterSelf(hyperlinkBase1);
                hLinkBase.Remove();
            }


            List<XElement> alldescendants = coreXDoc.Descendants().ToList();

            XName creator = dc + "creator";
            XName title = dc + "title";
            XName subject = dc + "subject";
            XName description = dc + "description";
            XName keywords = cp + "keywords";
            XName category = cp + "category";
            XName contentStatus = cp + "contentStatus";
            coreXDoc.Descendants()
                .Where(x => x.Name == creator || x.Name == title || x.Name == subject
                    || x.Name == description || x.Name == keywords || x.Name == category || x.Name == contentStatus)
                .Remove();
        } // End doRemoveBuiltInProperties


        private static void doUncheckCheckboxs(WordprocessingDocument wDoc)
        {
            MainDocumentPart mainPart = wDoc.MainDocumentPart;

            List<OpenXmlElement> ParagraphsWithBookmarks= mainPart.Document.Body.Descendants()
              .Where(a => a.GetType().Name.ToString() == "Paragraph"
              && (a.Descendants().Where(b => b.GetType().Name.ToString() == "BookmarkStart")).ToList().Count > 0).ToList();

            
            List<OpenXmlElement> cellsWithBookmarks = mainPart.Document.Body.Descendants()
              .Where(a => a.GetType().Name.ToString() == "TableCell"
              && (a.Descendants().Where(b => b.GetType().Name.ToString() == "BookmarkStart")).ToList().Count > 0).ToList();


            string bookmarkID = ParagraphsWithBookmarks.First().Descendants()
                                    .Where(a => a.GetType().Name.ToString() == "BookmarkStart").First()
                                    .GetAttributes().Where(a => a.LocalName.ToString() == "id").FirstOrDefault().Value.ToString();

            //CheckBox CheckboxInCurrentCellorParagraph = ParagraphsWithBookmarks.First().Descendants()
            //                        .Where(a => a.GetType().Name.ToString() == "BookmarkStart").First()
            //                        .GetAttributes().Where(a => a.LocalName.ToString() == "id").FirstOrDefault().Value.ToString();


            //   && a.Descendants().Where(b => b.LocalName == "rPr").FirstOrDefault() != null).ToList();
            foreach (OpenXmlElement element in ParagraphsWithBookmarks)
            {
                Console.WriteLine("____________________________");
                Console.WriteLine(element.GetType());
                Console.WriteLine(element.LocalName);
                Console.WriteLine(element.XName);
                Console.WriteLine("____________________________");
                
            }

        }




            #endregion

            #region Util Methods

            // area for helper functions 

            public static WordprocessingDocument getWordprocessingDocument(string docpath)
        {
            WordprocessingDocument wdDoc = null;
            try
            {
                wdDoc = WordprocessingDocument.Open(docpath, true);

            }
            catch (OpenXmlPackageException e)
            {

                if (e.ToString().ToLower().Contains("malformed uri"))
                {
                    Uri brokenUri = new Uri("http://broken-link/");

                    XNamespace relNs = "http://schemas.openxmlformats.org/package/2006/relationships";
                    using (FileStream fs = new FileStream(docpath, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                    {


                        using (ZipArchive za = new ZipArchive(fs, ZipArchiveMode.Update))
                        {
                            foreach (var entry in za.Entries.ToList())
                            {
                                if (!entry.Name.EndsWith(".rels"))
                                    continue;
                                bool replaceEntry = false;
                                XDocument entryXDoc = null;
                                using (var entryStream = entry.Open())
                                {
                                    try
                                    {
                                        entryXDoc = XDocument.Load(entryStream);
                                        if (entryXDoc.Root != null && entryXDoc.Root.Name.Namespace == relNs)
                                        {
                                            var urisToCheck = entryXDoc
                                                .Descendants(relNs + "Relationship")
                                                .Where(r => r.Attribute("TargetMode") != null && (string)r.Attribute("TargetMode") == "External");
                                            foreach (var rel in urisToCheck)
                                            {
                                                var target = (string)rel.Attribute("Target");
                                                if (target != null)
                                                {
                                                    try
                                                    {
                                                        Uri uri = new Uri(target);
                                                    }
                                                    catch (UriFormatException)
                                                    {
                                                        Uri newUri = brokenUri;
                                                        rel.Attribute("Target").Value = newUri.ToString();
                                                        replaceEntry = true;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    catch (XmlException)
                                    {
                                        continue;
                                    }
                                }
                                if (replaceEntry)
                                {
                                    var fullName = entry.FullName;
                                    entry.Delete();
                                    var newEntry = za.CreateEntry(fullName);
                                    using (StreamWriter writer = new StreamWriter(newEntry.Open()))
                                    using (XmlWriter xmlWriter = XmlWriter.Create(writer))
                                    {
                                        entryXDoc.WriteTo(xmlWriter);
                                    }
                                }
                            }
                        }
                    }
                    wdDoc = WordprocessingDocument.Open(docpath, true);
                }
            }
            return wdDoc;
        }
            #endregion
        
    }
}


public class MissingFilePropertiesException : Exception
{
    public MissingFilePropertiesException(string doctype)
       : base(string.Format("This document is missing default properties and is not a properly formed XML document. Please open the document in {0}, perform a Save As, and reprocess the new document.", doctype))
    {
    }
}