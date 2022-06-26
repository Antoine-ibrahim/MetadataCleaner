using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using System.Xml.Linq;
//using OpenXmlPowerTools;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.ExtendedProperties;
using System.Configuration;
using System.IO.Compression;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;

using A = DocumentFormat.OpenXml.Drawing;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using Sht = DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Office2010.Excel;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;
namespace MetadataCleaner
{
    public class ExcelCleaner
    {

        #region Static Variables

        private static readonly List<string> SPREASHEET_EXTENSIONS = new List<string> { ".xlsx", ".xlsm", ".xltx", ".xltm", ".xlsb", ".xlam" };

        private static readonly XNamespace cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
        private static readonly XNamespace dc = "http://purl.org/dc/elements/1.1/";
        private static readonly XNamespace dcterms = "http://purl.org/dc/terms/";
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

            if (SPREASHEET_EXTENSIONS.Contains(extension))
            {
                using (SpreadsheetDocument ExcelDoc = getSpreadsheetDocument(pathToDocument))
                {
                    try
                    {
                        // Get XDocument to perform direct XML edit operations with.
                        XDocument coreXDoc = ExcelDoc.CoreFilePropertiesPart.GetXDocument();

                        doRemoveSystemProperties(ExcelDoc, coreXDoc);      //done
                        doRemoveBuiltInProperties(ExcelDoc, coreXDoc);     //done
                        doRemoveCustomProperties(ExcelDoc);                //done
                        doRemoveTrackedChanges(ExcelDoc);                 //DONE NEW
                        doRemoveComments(ExcelDoc);                        //done
                      //doReplaceEmbedsWithPictures(ExcelDoc);             
                        doRemoveHyperlinks(ExcelDoc);                       //done
                        doRemoveDefinedNames(ExcelDoc);                     // done 
                        doRemovePersonalInfo(ExcelDoc, coreXDoc);           //done                 
                        doRemoveScinarios(ExcelDoc);                        //done 
                      //doRemoveHiddenRows(ExcelDoc);                     //done         // this  function is commented because it is dangerous, and Payne only "unhides" hidden rows instead of deleting them 
                        doUnhideHiddenRows(ExcelDoc);                       // added this function to mimic payne's functionality for hidden rows (instead of removing them)
                        doRemoveHiddenColumns(ExcelDoc);                    //Done NEW      // makes them visible instead of hidden
                      //doRemoveHiddenWorksheets(ExcelDoc);                 // commented due to function being dangerous, should be changed to unhide hidden worksheets instead
                        doRemoveHiddenObjects(ExcelDoc);                    //done  
                      //doRemoveLinksToExternalFiles(ExcelDoc);           
                        doRemoveCustomViews(ExcelDoc);                    //done               
                        doRemoveAutoFilter(ExcelDoc);                     // done , if there are hidden rows , they will be revealed using this function 
                        doRemoveCustomStyle(ExcelDoc);                     // DONE NEW  
                        doRemoveSparkLines(ExcelDoc);                      //done  
                        doRemoveHeaderAndFooter(ExcelDoc);                  //done
                        doRemoveSlicer(ExcelDoc);                           //done
                                                                            //AnalysePivotTableCache(ExcelDoc);???              //


                        // Save the document XML back to their document parts.
                        coreXDoc.Save(ExcelDoc.CoreFilePropertiesPart.GetStream(FileMode.Create, FileAccess.Write));
                    }
                    catch (Exception ex)
                    {
                        if (ex.Message.Contains("System.OutOfMemoryException"))
                        {
                            throw new FilePossiblyTooBiGException(ex.Message);
                        }
                        if (ex.GetType().Name == "MissingFilePropertiesException")
                        {
                            throw new MissingFilePropertiesException("Excel");
                        }
                        else
                        {
                            throw new Exception(ex.Message);
                        }
                    }
                }
            }
        }// End Clean() method. 




        //==== RemoveSystemProperties ===================================================
        /// <summary>
        /// This function removes all properties that are recorded by the system automatically, such as dates, authors, last modified etc...  
        /// </summary>
        /// <param name="path">This is the full path of the document that we are trying to remove Metadata From.</param>
        //=========================================================================
        public static void RemoveSystemProperties(string path)
        {
            using (SpreadsheetDocument ExcelDoc = getSpreadsheetDocument(path))
            {
                try
                {
                    XDocument coreXDoc = ExcelDoc.CoreFilePropertiesPart.GetXDocument();

                    doRemoveSystemProperties(ExcelDoc, coreXDoc);

                    coreXDoc.Save(ExcelDoc.CoreFilePropertiesPart.GetStream(FileMode.Create, FileAccess.Write));
                }
                catch (Exception ex)
                {

                    if (ex.Message.Contains("System.OutOfMemoryException"))
                    {
                        throw new FilePossiblyTooBiGException(ex.Message);
                    }
                    else {
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


            using (SpreadsheetDocument ExcelDoc = getSpreadsheetDocument(path))
            {
                try
                {
                    XDocument coreXDoc = ExcelDoc.CoreFilePropertiesPart.GetXDocument();

                    doRemoveBuiltInProperties(ExcelDoc, coreXDoc);

                    coreXDoc.Save(ExcelDoc.CoreFilePropertiesPart.GetStream(FileMode.Create, FileAccess.Write));
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

        } // End RemoveBuiltInProperties



        //==== RemoveCustomProperties ===================================================
        /// <summary>
        /// Removes all custom properties that were added to the document by users.   
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================

        public static void RemoveCustomProperties(string path)
        {
            using (SpreadsheetDocument ExcelDoc = getSpreadsheetDocument(path))
            {
                try
                {
                    doRemoveCustomProperties(ExcelDoc);
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




        //==== RemovePersonalInfo ===================================================
        /// <summary>
        /// Removes all personal info from the document and adds a setting that automatically removes personal information if the file is edited in the future.
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================

        public static void RemovePersonalInfo(string path)
        {
            using (SpreadsheetDocument ExcelDoc = getSpreadsheetDocument(path))
            {
                try
                {
                    XDocument coreXDoc = ExcelDoc.CoreFilePropertiesPart.GetXDocument();
                    doRemovePersonalInfo(ExcelDoc, coreXDoc);

                    coreXDoc.Save(ExcelDoc.CoreFilePropertiesPart.GetStream(FileMode.Create, FileAccess.Write));
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
        } // End RemovePersonalInfo




        //==== RemoveComments ===================================================
        /// <summary>
        /// Removes all comments from the document and removes all references to the authors of the comments.
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================
        public static void RemoveComments(string path)
        {
            using (SpreadsheetDocument ExcelDoc = getSpreadsheetDocument(path))
            {
                try
                {
                    doRemoveComments(ExcelDoc);

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
        } // End RemoveCommentsDoc




        //public static void RemoveHiddenRows(string path)
        //{
        //    using (SpreadsheetDocument ExcelDoc = getSpreadsheetDocument(path))
        //    {
        //        try
        //        {
        //            doRemoveHiddenRows(ExcelDoc);
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
        //} // End RemoveHiddenRows



        //==== UnhideHiddenRows ===================================================
        /// <summary>
        /// Sets hidden rows to visible.
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================

        public static void UnhideHiddenRows(string path)
        {
            using (SpreadsheetDocument ExcelDoc = getSpreadsheetDocument(path))
            {
                try
                {
                    doUnhideHiddenRows(ExcelDoc);
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
        } // End RemoveHiddenRows



        //==== RemoveHyperlinks ===================================================
        /// <summary>
        /// Remove any hyperlinks or on click links. This includes items such as a hyperlinked image.
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================

        public static void RemoveHyperlinks(string path)
        {
            using (SpreadsheetDocument ExcelDoc = getSpreadsheetDocument(path))
            {
                try
                {
                    doRemoveHyperlinks(ExcelDoc);
                }
                catch (Exception ex)
                {
                    if(ex.Message.Contains("System.OutOfMemoryException"))
                        throw new FilePossiblyTooBiGException(ex.Message);
                }

            }
        } // End RemoveHyperlinks




        //==== RemoveHiddenObjects ===================================================
        /// <summary>
        /// This removes all hidden objects in the document, for example hidden pictures or text boxes.   
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================
        public static void RemoveHiddenObjects(string path)
        {
            using (SpreadsheetDocument ExcelDoc = getSpreadsheetDocument(path))
            {
                try
                {
                    doRemoveHiddenObjects(ExcelDoc);
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




        //==== RemoveScinarios ===================================================
        /// <summary>
        /// This removes all Scinarios from the excel document.
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================

        public static void RemoveScinarios(string path)
        {
            using (SpreadsheetDocument ExcelDoc = getSpreadsheetDocument(path))
            {
                try
                {
                    doRemoveScinarios(ExcelDoc);

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


        }// End Remove Scinarios


        //==== RemoveCustomViews ===================================================
        /// <summary>
        /// This removes all custom views from the excel document.
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================


        public static void RemoveCustomViews(string path)
        {
            using (SpreadsheetDocument ExcelDoc = getSpreadsheetDocument(path))
            {
                try
                {
                    doRemoveCustomViews(ExcelDoc);
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
        }// End RemoveCustomView



        //==== RemoveDefinedNames ===================================================
        /// <summary>
        /// This removes all Devined Names in the excel document and updates the formulas that reference them with the correct values.
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================

        public static void RemoveDefinedNames(string path)
        {

            using (SpreadsheetDocument ExcelDoc = getSpreadsheetDocument(path))
            {
                try
                {
                    doRemoveDefinedNames(ExcelDoc);
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

        }// End RemoveDefinedNames





        //==== RemoveAutoFilter ===================================================
        /// <summary>
        /// This removes all auto filters in the excel document and unhides hidden rows.
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================
        public static void RemoveAutoFilter(string path)
        {
            using (SpreadsheetDocument ExcelDoc = getSpreadsheetDocument(path))
            {
                try
                {
                    doRemoveAutoFilter(ExcelDoc);
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
        }// End RemoveAutoFilter




        //==== RemoveHeaderAndFooter ===================================================
        /// <summary>
        /// This removes the header and footer from the excel document.
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================
        public static void RemoveHeaderAndFooter(string path)
        {
            using (SpreadsheetDocument ExcelDoc = getSpreadsheetDocument(path))
            {
                try
                {
                    doRemoveHeaderAndFooter(ExcelDoc);
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


        } //End RemoveHeaderAndFooter;


        //==== RemoveSlicer ===================================================
        /// <summary>
        /// This removes all slicers from the excel document.
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================

        public static void RemoveSlicer(string path)
        {
            using (SpreadsheetDocument ExcelDoc = getSpreadsheetDocument(path))
            {
                try
                {
                    doRemoveSlicer(ExcelDoc);
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

        }// End RemoveSlicer



        //==== RemoveSparkLines ===================================================
        /// <summary>
        /// This removes all Spark Lines from the excel document.
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================

        public static void RemoveSparkLines(string path)
        {
            using (SpreadsheetDocument ExcelDoc = getSpreadsheetDocument(path))
            {
                try
                {
                    doRemoveSparkLines(ExcelDoc);
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

        } //End RemoveSparkLines




        //==== RemoveCustomStyle ===================================================
        /// <summary>
        /// This removes all custom style added to the excel document and leaves the document with just the default custom style. 
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================

        public static void RemoveCustomStyle(string path)
        {
            using (SpreadsheetDocument ExcelDoc = getSpreadsheetDocument(path))
            {
                try
                {
                    doRemoveCustomStyle(ExcelDoc);
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
        } //End RemoveCustomStyle


        //public static void ReplaceEmbedsWithPictures(string path)
        //{
        //    using (SpreadsheetDocument ExcelDoc = getSpreadsheetDocument(path))
        //    {
        //        doReplaceEmbedsWithPictures(ExcelDoc);  
        //    }

        //}// End  //ReplaceEmbedsWithPictures





        //==== RemoveHiddenColumns ===================================================
        /// <summary>
        /// Set any hidden columns in the excel document to be visible.
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================

        public static void RemoveHiddenColumns(string path)
        {
            using (SpreadsheetDocument ExcelDoc = getSpreadsheetDocument(path))
            {
                try
                {
                    doRemoveHiddenColumns(ExcelDoc);
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
        }// End RemoveHiddenColumns

        //public static void RemoveHiddenSheets(string path)
        //{
        //    using (SpreadsheetDocument ExcelDoc = getSpreadsheetDocument(path))
        //    {
        //        try
        //        {
        //            doRemoveHiddenSheets(ExcelDoc);
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
        //}// End RemoveHiddenColumns




        //==== RemoveTrackedChanges ===================================================
        /// <summary>
        /// Approves all revisions, including deleted items, and edited items. If revision tracking is 
        /// on, it will be turned off 
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================
        public static void RemoveTrackedChanges(string path)
        {
            using (SpreadsheetDocument ExcelDoc = getSpreadsheetDocument(path))
            {
                try
                {
                    doRemoveTrackedChanges(ExcelDoc);
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
        }// End RemoveHiddenColumns


        
        //public static void RemoveLinksToExternalFiles(string path)
        //{
        //    using (SpreadsheetDocument ExcelDoc = getSpreadsheetDocument(path))
        //    {
        //        try
        //        {

        //            doRemoveLinksToExternalFiles(ExcelDoc);
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
        //}// End RemoveLinksToExternalFiles

        #endregion

        #region Private Methods
        private static void doRemoveSystemProperties(SpreadsheetDocument ExcelDoc, XDocument coreXDoc)
        {

            if (coreXDoc.Descendants().Count() <= 2 || ExcelDoc.ExtendedFilePropertiesPart.Properties.Count() <= 1)
            {
                throw new MissingFilePropertiesException("Excel");
            }

            WorkbookPart workbookPart = ExcelDoc.WorkbookPart;
            ExtendedFilePropertiesPart docProperties = ExcelDoc.ExtendedFilePropertiesPart;

            // change the appversion to 12.0000
            OpenXmlElement AppVersion = docProperties.Properties
                .Where(a => a.LocalName == "AppVersion").FirstOrDefault();
            Ap.ApplicationVersion applicationVersion1 = new Ap.ApplicationVersion();
            applicationVersion1.Text = "16.0300";
            AppVersion.InsertAfterSelf(applicationVersion1);
            AppVersion.Remove();


            // edit the created and modified date to the date that the cleaner was run.
            XName created = dcterms + "created";
            XName modified = dcterms + "modified";
            XName printed = cp + "lastPrinted";


             coreXDoc.Descendants()
                .Where(x => x.Name == created || x.Name == modified || x.Name == printed)
                .ToList()
                .ForEach(e => e.Value = DateTime.Now.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'"));


            OpenXmlElement totalTimes = docProperties.Properties
                .Where(a => a.LocalName == "TotalTime").FirstOrDefault();

            Ap.TotalTime newTotalTimes = new Ap.TotalTime();
            newTotalTimes.Text = "0";

            //   IEnumerable<OpenXmlElement> CorePropertiesToDelete;
            if (totalTimes != null)
            {
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


        } // End doRemoveSystemProperties


        private static void doRemoveBuiltInProperties(SpreadsheetDocument ExcelDoc, XDocument coreXDoc)
        {

            if (coreXDoc.Descendants().Count() <= 2 || ExcelDoc.ExtendedFilePropertiesPart.Properties.Count() <= 1)
            {
                throw new MissingFilePropertiesException("Excel");
            }

            WorkbookPart workbookPart = ExcelDoc.WorkbookPart;
            ExtendedFilePropertiesPart docProperties = ExcelDoc.ExtendedFilePropertiesPart;

            OpenXmlElement company = docProperties.Properties
                .Where(a => a.LocalName == "Company").FirstOrDefault();

            OpenXmlElement hLinkBase = docProperties.Properties
                .Where(a => a.LocalName == "HyperlinkBase").FirstOrDefault();

            if (company != null)
            {
                Ap.Company company1 = new Ap.Company();
                company.InsertAfterSelf(company1);
                company.Remove();

            }

            if (hLinkBase != null)
            {
                Ap.HyperlinkBase hyperlinkBase1 = new Ap.HyperlinkBase();
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
                 .ToList().ForEach(e => e.Value = DateTime.Now.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'"));

        } // End doRemoveBuiltInProperties




        private static void doRemoveCustomProperties(SpreadsheetDocument ExcelDoc)
        {
            // removes custom Properties
            CustomFilePropertiesPart CustomProperties = ExcelDoc.CustomFilePropertiesPart;
            ExcelDoc.DeletePart(CustomProperties);
        } // End doRemoveCustomProperties




        private static void doRemovePersonalInfo(SpreadsheetDocument ExcelDoc, XDocument coreXDoc)
        {

            WorkbookPart workbookPart = ExcelDoc.WorkbookPart;
            ExtendedFilePropertiesPart ExtendedProperties = ExcelDoc.ExtendedFilePropertiesPart;

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

            // Remove all Comment Authors and replaces it with just the default "Author".
            List<WorksheetPart> Worksheets = ExcelDoc.WorkbookPart.WorksheetParts.Where(a => a.GetType().Name == "WorksheetPart").ToList();

            foreach (WorksheetPart worksheet in Worksheets)
            {

                WorksheetCommentsPart commentsPart = worksheet.WorksheetCommentsPart;
                if (commentsPart != null) {
                    // remove all authors from the author section and replaces it with the default "Author"
                    OpenXmlElement authorPart = commentsPart.Comments.Authors;
                    Sht.Authors authors = new Sht.Authors();
                    Sht.Author author = new Sht.Author();
                    author.Text = "Author";
                    authorPart.InsertAfterSelf(author);
                    authorPart.Remove();

                    // coes through every comment in the comment list and updates the Text for the author name to specify the default author
                    OpenXmlElementList commentsList = commentsPart.Comments.CommentList.ChildElements;
                    foreach (OpenXmlElement comment in commentsList)
                    {
                        OpenXmlElement AuthorContainer = comment.Descendants().Where(a => a.GetType().Name == "Run").FirstOrDefault().ChildElements.Where(a => a.GetType().Name == "Text").FirstOrDefault();
                        Sht.Text text = new Sht.Text();
                        text.Text = "Author:";

                        AuthorContainer.InsertAfterSelf(text);
                        AuthorContainer.Remove();

                    }
                }
            }
        } // End doRemovePersonalInfo



        private static void doRemoveComments(SpreadsheetDocument ExcelDoc)
        {
            // retrieve a list of all the workbooks, because each one has its own comments section 
            WorkbookPart workbookPart = ExcelDoc.WorkbookPart;
            List<WorksheetPart> Worksheets = workbookPart.WorksheetParts.Where(a => a.GetType().Name == "WorksheetPart").ToList();

            // get and delete the comments part
            foreach (WorksheetPart worksheet in Worksheets)
            {

                WorksheetCommentsPart commentsPart = worksheet.WorksheetCommentsPart;
                worksheet.DeletePart(commentsPart);
                
            }
        }// End doRemoveCommentsDoc


        
        //this function may need to be revisited to prevent the deletion of hidden rows if the sheet has filtering in it
        // filtering may casuse unexpected rows to be deleted. 
        private static void doRemoveHiddenRows(SpreadsheetDocument ExcelDoc) {
            WorkbookPart workbookPart = ExcelDoc.WorkbookPart;
            List<WorksheetPart> worksheets = workbookPart.WorksheetParts.Where(a => a.GetType().Name == "WorksheetPart").ToList();
            Sht.SharedStringTable sharedStringTable = workbookPart.SharedStringTablePart.SharedStringTable;

            // check every sheet to see if it has any hidden rows 
            foreach (WorksheetPart worksheet in worksheets)
            {
                List<OpenXmlElement> HiddenRows = worksheet.Worksheet.Descendants().Where(a => a.GetType().Name == "Row"  && a.GetAttributes().Where(b => b.LocalName == "hidden").FirstOrDefault().Value == "1").ToList();
                foreach (OpenXmlElement row in HiddenRows)
                {

                   // if a hidden row is found, we will need to iterate through every single cell in that row and see what item it is referencing in the sharedstring 
                   List< OpenXmlElement> cells = row.ChildElements.Where(a => a.GetType().Name == "Cell").ToList();
                    foreach (Sht.Cell cell in cells)
                    {
                        
                        string val = cell.InnerText;
                        if (cell.DataType != null && cell.DataType.Value.ToString() == "SharedString") {
                            // after finding what item it is referencing in the shared string we will need to go through every other sheet/row/column to see if there is any 
                            // other items also referencing this same cell, 
                            if (!CheckAllColumnsforSameReference(ExcelDoc, val))
                            {

                                // if there is only one , we need to remove the element from the shared string 
                                string tbd = sharedStringTable.ElementAt(Convert.ToInt32(val)).InnerText;
                                sharedStringTable.ElementAt(Convert.ToInt32(val)).Remove();

                                // and update every column that has a number larger then the reference , decrement it by 1, so it is now referencing the correct item
                                editAllCelsWIthCorrectRefs(ExcelDoc, val);


                            }
                        }
                        cell.Remove();
                    }
                    row.Remove();
                }
            }
        }


        private static void doUnhideHiddenRows(SpreadsheetDocument ExcelDoc)
        {

            WorkbookPart workbookPart = ExcelDoc.WorkbookPart;
            List<WorksheetPart> worksheets = workbookPart.WorksheetParts.Where(a => a.GetType().Name == "WorksheetPart").ToList();
            Sht.SharedStringTable sharedStringTable = workbookPart.SharedStringTablePart.SharedStringTable;

            // check every sheet to see if it has any hidden rows 
            foreach (WorksheetPart worksheet in worksheets)
            {
                List<OpenXmlElement> HiddenRows = worksheet.Worksheet.Descendants().Where(a => a.GetType().Name == "Row" && a.GetAttributes().Where(b => b.LocalName == "hidden").FirstOrDefault().Value == "1").ToList();
                foreach (OpenXmlElement row in HiddenRows)
                {
                    List<OpenXmlAttribute> allatts = row.GetAttributes().ToList();
                    
                    OpenXmlAttribute hidden= row.GetAttributes().Where(a => a.LocalName == "hidden").FirstOrDefault();
                    OpenXmlAttribute dyDescent = row.GetAttributes().Where(a => a.LocalName == "dyDescent").FirstOrDefault();

                    if (hidden != null && hidden.Value != "") {
                        row.RemoveAttribute(hidden.LocalName, hidden.NamespaceUri);
                    }

                    if (dyDescent != null && dyDescent.Value != "") {
                        row.RemoveAttribute(dyDescent.LocalName, dyDescent.NamespaceUri);
                    }
                }
            }

        }

        private static void doRemoveHyperlinks(SpreadsheetDocument ExcelDoc)
        {
            WorkbookPart workbookPart = ExcelDoc.WorkbookPart;
            List<WorksheetPart> worksheets = workbookPart.WorksheetParts.Where(a => a.GetType().Name == "WorksheetPart").ToList();
            foreach (WorksheetPart worksheet in worksheets)
            {
                //get all hyperlinks in each slide and remove it and its references 

                OpenXmlElement HyperlinkContainer = worksheet.Worksheet.Descendants().Where(a => a.GetType().Name == "Hyperlinks").FirstOrDefault();
                if (HyperlinkContainer != null)
                {
                    List<OpenXmlElement> hyperlinks = HyperlinkContainer.Descendants().Where(a => a.GetType().Name == "Hyperlink").ToList();

                    foreach (OpenXmlElement hyperlink in hyperlinks)
                    {
                        string id = hyperlink.GetAttributes()
                            .Where(a => a.LocalName == "id").FirstOrDefault().Value;
                        if (id != null)
                        {
                            ReferenceRelationship hlr = worksheet.GetReferenceRelationship(id);

                            
                            worksheet.DeleteReferenceRelationship(hlr);
                        }
                        hyperlink.Remove();

                    }
                    HyperlinkContainer.Remove();

                    //for hyperlinked pictures , check if the excell sheet has any drawing parts and check if they have hyperlinks 

                    if (worksheet.DrawingsPart != null)
                    {
                        List<OpenXmlElement> DrawingsHLinks = worksheet.DrawingsPart.WorksheetDrawing.Descendants().Where(a => a.LocalName == "hlinkClick").ToList();
                        foreach (OpenXmlElement HlinkElement in DrawingsHLinks)
                        {

                            string id = HlinkElement.GetAttributes()
                               .Where(a => a.LocalName == "id").FirstOrDefault().Value;

                            ReferenceRelationship hlr = worksheet.DrawingsPart.GetReferenceRelationship(id);
                            // ReferenceRelationship hlr = worksheet.GetReferenceRelationship(id);

                            HlinkElement.Remove();
                            worksheet.DrawingsPart.DeleteReferenceRelationship(hlr);
                        }
                    }

                }

            }
        }

        private static void doRemoveHiddenObjects(SpreadsheetDocument ExcelDoc)
        {
            WorkbookPart workbookPart = ExcelDoc.WorkbookPart;
            List<WorksheetPart> worksheets = workbookPart.WorksheetParts.Where(a => a.GetType().Name == "WorksheetPart").ToList();
            foreach (WorksheetPart worksheet in worksheets)
            {
                if (worksheet.DrawingsPart != null)
                {
                    List<OpenXmlElement> HiddenObjects = worksheet.DrawingsPart.WorksheetDrawing.Descendants()
                        .Where(a => a.GetType().Name == "NonVisualDrawingProperties")
                        .Where(a => a.GetAttributes().Where(b => b.LocalName == "hidden").FirstOrDefault().Value == "1").ToList();

                    foreach (OpenXmlElement HiddenObject in HiddenObjects)
                    {
                        try
                        {
                            OpenXmlElement objContainer = HiddenObject.Ancestors().Where(a => a.GetType().Name == "TwoCellAnchor").FirstOrDefault();
                            objContainer.Remove();

                        }
                        catch (Exception )
                        {
                            //Console.WriteLine("This document has a hidden object with an unknown container type check the drawings.xml for this document to investigate");
                        }
                        
                    }
                }
            }
        }


        private static void doRemoveScinarios(SpreadsheetDocument ExcelDoc)
        {
            WorkbookPart workbookPart = ExcelDoc.WorkbookPart;
            List<WorksheetPart> worksheets = workbookPart.WorksheetParts.Where(a => a.GetType().Name == "WorksheetPart").ToList();
            foreach (WorksheetPart worksheet in worksheets)
            {
                List<OpenXmlElement> Scenarios = worksheet.Worksheet.Descendants().Where(a => a.GetType().Name == "Scenarios").ToList();
                foreach (OpenXmlElement scenario in Scenarios)
                {
                    scenario.Remove();
                }
            }

        }


        private static void doRemoveCustomViews(SpreadsheetDocument ExcelDoc)
        {
            WorkbookPart workbookPart = ExcelDoc.WorkbookPart;
            List<WorksheetPart> worksheets = workbookPart.WorksheetParts.Where(a => a.GetType().Name == "WorksheetPart").ToList();
            foreach (WorksheetPart worksheet in worksheets)
            {
                List<OpenXmlElement> worksheetViews = worksheet.Worksheet.Descendants().Where(a => a.GetType().Name == "CustomSheetViews").ToList();
                foreach (OpenXmlElement customView in worksheetViews)
                {
                    customView.Remove();
                }
            }

            OpenXmlElement workbookViews = workbookPart.Workbook.CustomWorkbookViews;
            if (workbookViews != null) {
                workbookViews.Remove();
            }
            
           
        }

        // this will remove the defined names and update all formulas by replacing the all defined variables with their values instead.
        private static void doRemoveDefinedNames(SpreadsheetDocument ExcelDoc)
        {
            WorkbookPart workbookPart = ExcelDoc.WorkbookPart;
            List<WorksheetPart> worksheets = workbookPart.WorksheetParts.Where(a => a.GetType().Name == "WorksheetPart").ToList();
            OpenXmlElement definedNamesContainer = workbookPart.Workbook.DefinedNames;

            List<string> delimiters = new List<string>() { "*","+","*","/", "%", "^", "(", ")", "&"};

            Dictionary<string, string> ListOfDefinedNames = new Dictionary<string, string>();        
            if (definedNamesContainer != null)
            {

                List<OpenXmlElement> definedNames = definedNamesContainer.ChildElements.ToList();
                foreach (OpenXmlElement definedName in definedNames)
                {
                    string Name = definedName.GetAttributes().Where(a => a.LocalName == "name").FirstOrDefault().Value;
                    string content = definedName.InnerText;
                    string sheetid = definedName.GetAttributes().Where(a => a.LocalName == "localSheetId").FirstOrDefault().Value;
                    if (!Name.Contains("_FilterDatabase") && sheetid ==null)
                    {
                        ListOfDefinedNames.Add(Name, content);
                    }
                    
                }
              
                
                // update all formulas
                foreach (WorksheetPart worksheet in worksheets)
                {
                    List<OpenXmlElement> cellFormulas = worksheet.Worksheet.Descendants().Where(a => a.GetType().Name == "CellFormula").ToList();
                    foreach (Sht.CellFormula formula in cellFormulas)
                    {
                        string equation = formula.Text;
                        string NewFormual = "";
                        string CurrentText = "";
                        foreach (char character in equation)
                        {                           
                            if (delimiters.Contains(character.ToString()))
                            {// if the character is a delimiter 
                                if (CurrentText == "")
                                {
                                    NewFormual += character;
                                }
                                else {
                                    try
                                    {
                                        string DefinedValue = ListOfDefinedNames[CurrentText];
                                        NewFormual += DefinedValue + character;
                                        CurrentText = "";
                                    }
                                    catch
                                    {
                                        NewFormual += CurrentText + character;
                                        CurrentText = "";
                                    }            
                                }
                            }


                            else {
                                // if the character is not a delimiter
                                CurrentText += character;
                            }
                        }

                        if (CurrentText != "")
                        {
                            NewFormual += CurrentText;
                            CurrentText = "";
                        }
                        // update the formual 
                        //Console.WriteLine("Old: " + formula.Text + "   New Formula: " + NewFormual);
                        formula.Text = NewFormual;
                    }
                }
                // remove the container with the defined names, after all formulas have been fixed to no longer include defined values
                definedNamesContainer.Remove();
            }
        }


        // may need to come back and edit this code , it clashes with the remove hidden row function because it has to unhide all hidden rows
        private static void doRemoveAutoFilter(SpreadsheetDocument ExcelDoc)
        {
            WorkbookPart workbookPart = ExcelDoc.WorkbookPart;
            List<WorksheetPart> worksheets = workbookPart.WorksheetParts.Where(a => a.GetType().Name == "WorksheetPart").ToList();
            foreach (WorksheetPart worksheet in worksheets)
            {
                List<OpenXmlElement> autoFilter = worksheet.Worksheet.Descendants().Where(a => a.GetType().Name == "AutoFilter").ToList();
                foreach (OpenXmlElement Filter in autoFilter)
                {
                    Filter.Remove();
                }
                if (worksheet.Worksheet.SheetProperties != null)
                {
                    worksheet.Worksheet.SheetProperties.Remove();
                }
                

                List<OpenXmlElement> HiddenRows = worksheet.Worksheet.Descendants().Where(a => a.GetType().Name == "Row" && a.GetAttributes().Where(b => b.LocalName == "hidden").FirstOrDefault().Value == "1").ToList();
                foreach (OpenXmlElement row in HiddenRows)
                {

                    OpenXmlAttribute hiddenProperty = row.GetAttributes().Where(a => a.LocalName == "hidden").FirstOrDefault();
                    row.RemoveAttribute(hiddenProperty.LocalName, hiddenProperty.NamespaceUri);
                }     
            }
        }


        private static void doRemoveHeaderAndFooter(SpreadsheetDocument ExcelDoc)
        {
            WorkbookPart workbookPart = ExcelDoc.WorkbookPart;
            List<WorksheetPart> worksheets = workbookPart.WorksheetParts.Where(a => a.GetType().Name == "WorksheetPart").ToList();
            foreach (WorksheetPart worksheet in worksheets)
            {               
                List<OpenXmlElement> HeaderFooters = worksheet.Worksheet.Descendants().Where(a => a.GetType().Name == "HeaderFooter").ToList();
                foreach (OpenXmlElement headerFooter in HeaderFooters)
                {
                    headerFooter.Remove();

                }
            }


        }

        private static void doRemoveSlicer(SpreadsheetDocument ExcelDoc)
        {

            WorkbookPart workbookPart = ExcelDoc.WorkbookPart;
            OpenXmlElement definedNamesContainer = workbookPart.Workbook.DefinedNames;
            List<WorksheetPart> worksheets = workbookPart.WorksheetParts.Where(a => a.GetType().Name == "WorksheetPart").ToList();
            foreach (WorksheetPart worksheet in worksheets)
            {
                // get the slicerlist and slicer references from the worksheet 
                List<OpenXmlElement> slicerList = worksheet.Worksheet.Descendants().Where(a => a.GetType().Name == "SlicerList").ToList();
                List<OpenXmlElement> slicers = worksheet.Worksheet.Descendants().Where(a => a.GetType().Name == "SlicerRef").ToList();
                foreach (OpenXmlElement slicer in slicers)
                {
                   
                    string SlicerId = slicer.GetAttributes().Where(a => a.LocalName == "id").FirstOrDefault().Value;
                    SlicersPart slicerPart = worksheet.GetPartById(SlicerId) as SlicersPart;


                    List<OpenXmlElement> slicersList = slicerPart.Slicers.ChildElements.ToList();
                    foreach (Slicer slicerCache in slicersList)
                    {
                        // for every slicer, remove the defined name portion for it from the workbook
                        string slicerDefinedName = slicerCache.Cache;
                        List<OpenXmlElement> definedNames = definedNamesContainer.ChildElements.Where(a => a.GetAttributes().Where(b => b.LocalName == "name").FirstOrDefault().Value == slicerDefinedName).ToList();
                        foreach (OpenXmlElement definedName in definedNames)
                        {
                            definedName.Remove(); 
                        }
                    }

                    if (definedNamesContainer != null && definedNamesContainer.ChildElements.Count == 0)
                    {
                        definedNamesContainer.Remove();
                    }                 
                    slicer.Remove(); 
                    worksheet.DeletePart(slicerPart); 
                }
                if (slicerList != null)
                {
                    foreach (OpenXmlElement item in slicerList)
                    {
                        item.Parent.Remove();
                    }
                }
                
               
              // remove all slicers from the drawing part of the document 
                if (worksheet.DrawingsPart != null)
                {
                    List<OpenXmlElement> DrawingsSlicers = worksheet.DrawingsPart.WorksheetDrawing.Descendants().Where(a => a.GetType().Name == "Slicer").ToList();
                    foreach (OpenXmlElement DrawingsSlicer in DrawingsSlicers)
                    {

                        OpenXmlElement DrawingSlicerContainer = DrawingsSlicer.Ancestors().Where(a => a.GetType().Name == "TwoCellAnchor").FirstOrDefault();
                        DrawingSlicerContainer.Remove();
                    }

                    if (worksheet.DrawingsPart.WorksheetDrawing.Descendants().Count() ==0)
                    {
                        string refid = worksheet.GetIdOfPart(worksheet.DrawingsPart).ToString();
                        OpenXmlElement worksheetDrawingPart = worksheet.Worksheet.Descendants().Where(a => a.GetAttributes().Where(b => b.LocalName == "id").FirstOrDefault().Value == refid).FirstOrDefault();
                        worksheetDrawingPart.Remove();
                        worksheet.DeletePart(worksheet.DrawingsPart);
                        //do something here
                    }
                }

                OpenXmlElement ExtentionsList = worksheet.Worksheet.Descendants().Where(a => a.GetType().Name == "WorksheetExtensionList").FirstOrDefault();
                if (ExtentionsList != null)
                {
                    if (ExtentionsList.ChildElements.Count() == 0)
                    {
                        ExtentionsList.Remove();
                    }
                }
            }

            // delete all the slicer caches
            OpenXmlElement slicerCacheContainers = workbookPart.Workbook.Descendants().Where(a => a.GetType().Name == "SlicerCaches").FirstOrDefault();
            if (slicerCacheContainers != null)
            {
                foreach (SlicerCache cache in slicerCacheContainers.ChildElements)
                {
                    string id = cache.Id;
                    OpenXmlPart slicerCachePart = workbookPart.GetPartById(id);
                    workbookPart.DeletePart(slicerCachePart);

                }
         
                    OpenXmlElement slicerCacheExtension = slicerCacheContainers.Ancestors().Where(a => a.GetType().Name == "WorkbookExtension").FirstOrDefault();
                    slicerCacheExtension.Remove();

                
            }
        }



        private static void doRemoveSparkLines(SpreadsheetDocument ExcelDoc) {
            WorkbookPart workbookPart = ExcelDoc.WorkbookPart;
            List<WorksheetPart> worksheets = workbookPart.WorksheetParts.Where(a => a.GetType().Name == "WorksheetPart").ToList();
            foreach (WorksheetPart worksheet in worksheets)
            {
                List<OpenXmlElement> SparklineGroups = worksheet.Worksheet.Descendants().Where(a => a.GetType().Name == "SparklineGroups").ToList();
                foreach (OpenXmlElement sparklineGroup in SparklineGroups)
                {
                    sparklineGroup.Parent.Remove();

                }


                List<OpenXmlElement> worksheetExtensionList = worksheet.Worksheet.Descendants().Where(a => a.GetType().Name == "WorksheetExtensionList").ToList();
                foreach (OpenXmlElement extentionList in worksheetExtensionList)
                {
                    if (extentionList.ChildElements.Count == 0)
                    {
                        extentionList.Remove();
                    }
                }


            }


        }


        private static void doRemoveCustomStyle(SpreadsheetDocument ExcelDoc)
        {
            List<string> defaultRowAttributes = new List<string>() { "r", "spans", "ht", "dyDescent" };
            List<string> defaultCellAttributes = new List<string>() { "r", "t" };
            WorkbookPart workbookPart = ExcelDoc.WorkbookPart;
            List<WorksheetPart> worksheets = workbookPart.WorksheetParts.Where(a => a.GetType().Name == "WorksheetPart").ToList();

            foreach (WorksheetPart worksheet in worksheets)
            {
                List<OpenXmlElement> Rows = worksheet.Worksheet.Descendants().Where(a => a.GetType().Name == "Row").ToList();
                foreach (Sht.Row row in Rows)
                {
                   List<OpenXmlAttribute> rowAttributes = row.GetAttributes().Where(a => !defaultRowAttributes.Contains(a.LocalName)).ToList();
                    foreach (OpenXmlAttribute attribute in rowAttributes)
                    {
                        row.RemoveAttribute(attribute.LocalName, attribute.NamespaceUri);
                    }                  
                }

                List<OpenXmlElement> cells = worksheet.Worksheet.Descendants().Where(a => a.GetType().Name == "Cell").ToList();
                foreach (OpenXmlElement cell in cells)
                {
                    List<OpenXmlAttribute> cellAttributes = cell.GetAttributes().Where(a => !defaultCellAttributes.Contains(a.LocalName)).ToList();
                    foreach (OpenXmlAttribute attribute in cellAttributes)
                    {
                       cell.RemoveAttribute(attribute.LocalName, attribute.NamespaceUri);

                    }
                }

                //update the collumn styles 
                List<OpenXmlElement> columns = worksheet.Worksheet.Descendants().Where(a => a.GetType().Name == "Column").ToList();
                foreach (Sht.Column column in columns)
                {
                    List<OpenXmlAttribute> ColStyle = column.GetAttributes().Where(a => a.LocalName == "style").ToList();
                    if(ColStyle!=null && ColStyle.Count!=0 && ColStyle.FirstOrDefault().Value != "1")
                    {
                        column.Style = (UInt32Value)1U;
                    }

                }
                
                      
            }

            WorkbookStylesPart stylePart = workbookPart.WorkbookStylesPart;

            List<OpenXmlElement> fonts = stylePart.Stylesheet.Fonts.ToList();
            for (int i = fonts.Count; i > 1; i--)
            {
                fonts[i-1].Remove();

            }

            List<OpenXmlElement> fills = stylePart.Stylesheet.Fills.ToList();
            for (int i = fills.Count; i > 1; i--)
            {
                fills[i-1].Remove();

            }
            List<OpenXmlElement> borders = stylePart.Stylesheet.Borders.ToList();

            for (int i = borders.Count; i > 1; i--)
            {
                borders[i-1].Remove();

            }

            List<OpenXmlElement> cellStyleFormat = stylePart.Stylesheet.Descendants()
               .Where(a => a.LocalName == "xf").ToList();
            foreach (Sht.CellFormat cellFormat in cellStyleFormat)
            {
                if (cellFormat.FillId != 0 | cellFormat.FontId != 0 || cellFormat.BorderId != 0)
                {
                    cellFormat.Remove();
                }

            }

            List<OpenXmlElement> cellStyles = stylePart.Stylesheet.CellStyles.ToList();
            foreach (Sht.CellStyle style in cellStyles)
            {
                if (style.FormatId !=0 ||style.Name != "Normal")
                {
                    style.Remove();
                }
            }



        }

        private static void doReplaceEmbedsWithPictures(SpreadsheetDocument ExcelDoc)
        {
            WorkbookPart workbookPart = ExcelDoc.WorkbookPart;
            List<WorksheetPart> worksheets = workbookPart.WorksheetParts.Where(a => a.GetType().Name == "WorksheetPart").ToList();

            foreach (WorksheetPart worksheet in worksheets)
            {

                // remove the vml drawing part
                List<OpenXmlElement> vmlDrawings = worksheet.Worksheet.Descendants().Where(a => a.GetType().Name == "LegacyDrawing").ToList();
                foreach (Sht.LegacyDrawing vml in vmlDrawings)
                {
                    string id = vml.Id;
                    OpenXmlPart vmlDrawingpart =worksheet.GetPartById(id);
                    worksheet.DeletePart(vmlDrawingpart);
                    vml.Remove();
                }

                // remove the ole objects 
                OpenXmlElement OleObjectsPart = worksheet.Worksheet.Descendants().Where(a => a.GetType().Name == "OleObjects").FirstOrDefault();
                if (OleObjectsPart != null)
                {
                    foreach (OpenXmlElement element in OleObjectsPart.ChildElements)
                    {
                        // id of the embeded object
                        string OleObjId = element.Descendants()
                            .Where(a => a.GetType().Name == "OleObject").FirstOrDefault()
                            .GetAttributes().Where(a => a.LocalName == "id").FirstOrDefault().Value;

                        // id of the referenced image (ex rid1 => image1.emf)
                        string imageid = element.Descendants()
                            .Where(a => a.LocalName == "objectPr").FirstOrDefault()
                            .GetAttributes().Where(a => a.LocalName == "id").FirstOrDefault().Value;


                        // id of the shapeid for the drawing xml 
                        string shapeId = element.Descendants()
                            .Where(a => a.GetType().Name == "OleObject").FirstOrDefault()
                            .GetAttributes().Where(a => a.LocalName == "shapeId").FirstOrDefault().Value;


                        OpenXmlElement fromMarker = element.Descendants().Where(a => a.LocalName == "from").FirstOrDefault().CloneNode(true);

                        OpenXmlElement toMarker = element.Descendants().Where(a => a.LocalName == "to").FirstOrDefault().CloneNode(true);


                        OpenXmlElement shape=  worksheet.DrawingsPart.WorksheetDrawing.Descendants().Where(a=>a.LocalName=="cNvPr" && a.GetAttributes().Where(b => b.LocalName == "id").FirstOrDefault().Value == shapeId).FirstOrDefault();
                        string shapeName = shape.GetAttributes().Where(a => a.LocalName == "name").FirstOrDefault().Value;


                        OpenXmlElement newPicture = GenerateDrawingPart(worksheet,fromMarker, toMarker, imageid, shapeId, shapeName);


                        OpenXmlElement shapeContainer = shape.Ancestors().Where(a => a.LocalName == "AlternateContent" ).FirstOrDefault();


                        worksheet.DeletePart(OleObjId);

                    }
                    OleObjectsPart.Remove();
                }



            }

        }


        // note there is no way to link a hidden column to the cells of that particular column, instead this function will set any hidden columns to be visible. 
        private static void doRemoveHiddenColumns(SpreadsheetDocument ExcelDoc)
        {
            WorkbookPart workbookPart = ExcelDoc.WorkbookPart;
            List<WorksheetPart> worksheets = workbookPart.WorksheetParts.Where(a => a.GetType().Name == "WorksheetPart").ToList();
            foreach (WorksheetPart worksheet in worksheets)
            {
                // get a list of all hidden columns
                List<OpenXmlElement> columnContainer = worksheet.Worksheet.Descendants().Where(a => a.GetType().Name == "Columns").ToList();
                List <OpenXmlElement> customColumns = worksheet.Worksheet.Descendants().Where(a => a.GetType().Name == "Column").ToList();
                foreach (Sht.Column column in customColumns)
                {
                    if (column.Hidden!= null && column.Hidden == true)
                    {
                        column.Remove();
                    }

                }
                foreach (OpenXmlElement container in columnContainer)
                {
                    if (container.ChildElements.Count < 1)
                    {
                        container.Remove();
                    }
                }
   
            }

        }
        // incomplete , need to revisit this 
        private static void doRemoveHiddenSheets(SpreadsheetDocument ExcelDoc)
        {
            WorkbookPart workbookPart = ExcelDoc.WorkbookPart;
            List<WorksheetPart> worksheets = workbookPart.WorksheetParts.Where(a => a.GetType().Name == "WorksheetPart").ToList();
            foreach (WorksheetPart worksheet in worksheets)
            {
                // get a list of all hidden columns
                string id = workbookPart.GetIdOfPart(worksheet);
                List<OpenXmlElement> columnContainer = worksheet.Worksheet.Descendants().Where(a => a.GetType().Name == "Columns").ToList();
                List<OpenXmlElement> customColumns = worksheet.Worksheet.Descendants().Where(a => a.GetType().Name == "Column").ToList();
              

            }

        }

        private static void doRemoveTrackedChanges(SpreadsheetDocument ExcelDoc)
        {
            WorkbookPart workbookPart = ExcelDoc.WorkbookPart;

            // delete all the resion logs, and the user data for the revisions
            WorkbookRevisionHeaderPart RevisionHeaderPart = ExcelDoc.WorkbookPart.WorkbookRevisionHeaderPart;
            WorkbookUserDataPart UserDataPart = ExcelDoc.WorkbookPart.WorkbookUserDataPart;
            workbookPart.DeletePart(UserDataPart);
            workbookPart.DeletePart(RevisionHeaderPart);

            // turn on the setting that removes change tracking. 
            workbookPart.Workbook.WorkbookProperties.FilterPrivacy = true;
        }


        private static void doRemoveLinksToExternalFiles(SpreadsheetDocument ExcelDoc)
        {
            WorkbookPart workbookPart = ExcelDoc.WorkbookPart;
            List<string> listOfSheetNames = new List<string>();
            bool equationHasExternalShtName = false;
            Dictionary<string, string> sheetIds = new Dictionary<string, string>();
            List<OpenXmlElement> sheets = workbookPart.Workbook.Sheets.ToList();
            
            // create a dictionary that links all the sheet reference ids to the sheet ids from the workbook, will be needed for deletion of formula references 
            foreach (Sht.Sheet sheet in sheets)
            {
                sheetIds.Add(sheet.Id, sheet.SheetId);

            }

            List<ExternalWorkbookPart> ExternalWorkbooksList = workbookPart.ExternalWorkbookParts.ToList();
            foreach (ExternalWorkbookPart ExternalWorkbook in ExternalWorkbooksList)
            {
                List<OpenXmlElement> sheetNames = ExternalWorkbook.ExternalLink.Descendants().Where(a => a.GetType().Name == "SheetName").ToList();
                foreach (Sht.SheetName name in sheetNames)
                {
                    listOfSheetNames.Add(name.Val);
                }
            }

            // go through the formula cells and see if it contains a sheet name if it does delete it , it will have a value field to take its place in the cell 
            List<WorksheetPart> worksheets = workbookPart.WorksheetParts.Where(a => a.GetType().Name == "WorksheetPart").ToList();
            foreach (WorksheetPart worksheet in worksheets)
            {
                string referenceId = workbookPart.GetIdOfPart(worksheet);
                List<OpenXmlElement> cellFormulas = worksheet.Worksheet.Descendants().Where(a => a.GetType().Name == "CellFormula").ToList();
                foreach (Sht.CellFormula formula in cellFormulas)
                {
                    Sht.Cell cell=formula.Ancestors().Where(a => a.GetType().Name == "Cell").FirstOrDefault() as Sht.Cell;
                    string cellReference = cell.CellReference;
                    string equation = formula.Text;
                    // check if the equation in thr formula contains one of the sheet names
                    foreach (string shtName in listOfSheetNames)
                    {
                        if (equation.Contains("]"+shtName+"!$") || equation.Contains("]" + shtName + @"'!$"))
                        {                        
                            equationHasExternalShtName = true;
                        }

                    }

                    if (equation.Contains("[") && equation.Contains("]") && equationHasExternalShtName == true)
                    {
                        
                        formula.Remove();
                        try
                        {
                            OpenXmlElement calcChain = workbookPart.CalculationChainPart.CalculationChain.ChildElements
                                .Where(a => a.GetAttributes().Where(b => b.LocalName == "r").FirstOrDefault().Value == cellReference
                                && a.GetAttributes().Where(c => c.LocalName == "i").FirstOrDefault().Value == sheetIds[referenceId]).FirstOrDefault();
                            calcChain.Remove();
                        }
                        catch { }
                        equationHasExternalShtName = false;
                    }
                }

            }

            // now delete all the external links 


            foreach (ExternalWorkbookPart xWorkbookPart in ExternalWorkbooksList)
            {
                workbookPart.DeletePart(xWorkbookPart);
            }
            
            if (workbookPart.Workbook.ExternalReferences != null) {
                workbookPart.Workbook.ExternalReferences.Remove();
            }
           


            // need to add code up there to create a dictionary of sheetid to id of sheet id (ex Rid2 => sheet id 1)
            // we then need to check the id of the worksheet we are working on , and every time we remove a formula we need to update the calcchain to remove that cell for that sheet id 




        }


        #endregion


        #region Private Helper Methods

        // this is a helper function used while deleting rows, it checks all other columns in all spread sheets to see if they exist in the document
        // this function is used to determine if the cell should be deleted from the shared string and if shifting references is required.  

        private static bool CheckAllColumnsforSameReference(SpreadsheetDocument ExcelDoc, string val)
        {
            WorkbookPart workbookPart = ExcelDoc.WorkbookPart;
            List<WorksheetPart> worksheets = workbookPart.WorksheetParts.Where(a => a.GetType().Name == "WorksheetPart").ToList();
            int count = 0;
            foreach (WorksheetPart worksheet in worksheets)
            {
                List<OpenXmlElement> CellsWithSameReference = worksheet.Worksheet.Descendants().Where(a => a.GetType().Name == "Cell" && a.InnerText == val).ToList();
                if (CellsWithSameReference != null && CellsWithSameReference.Count > 0)
                {
                    count += CellsWithSameReference.Count;

                }

            }

            return count > 1;
        }

        // after deleting a column/row from the spread sheet, if there was an item removed from the shared string, then all references have to change in order to point
        // to the correct word/phrase in the shared string list 

        private static void editAllCelsWIthCorrectRefs(SpreadsheetDocument ExcelDoc, string val)
        {
            WorkbookPart workbookPart = ExcelDoc.WorkbookPart;
            List<WorksheetPart> worksheets = workbookPart.WorksheetParts.Where(a => a.GetType().Name == "WorksheetPart").ToList();
            //check each worksheet 
     
            foreach (WorksheetPart worksheet in worksheets)
            {
                //get list of every single cell that has a higher reference

                List <OpenXmlElement> cells = worksheet.Worksheet.Descendants()
                  .Where(a => a.GetType().Name == "Cell" &&
                  (a.GetAttributes().Where(b => b.LocalName == "t").FirstOrDefault().Value == "s")
                  && a.InnerText != ""
                  && Convert.ToInt32(a.InnerText) > Convert.ToInt32(val)).ToList();
                
                foreach (Sht.Cell cell in cells)
                {
                    string cellReference = cell.CellReference;
                    
                    int cellValue = Convert.ToInt32(cell.InnerText);
                    // string cellReference = cell.GetAttributes().Where(a => a.GetType().Name == "CellReference").FirstOrDefault().Value;

                    cell.CellValue.Text = (cellValue -1).ToString();
                }                
            }
        }

        public static void Test(string path)
        {
            using (SpreadsheetDocument ExcelDoc = SpreadsheetDocument.Open(path, true))
            {
                bool has = CheckAllColumnsforSameReference(ExcelDoc, "3");
               // Console.WriteLine(has);
            }
        }


        private static SpreadsheetDocument getSpreadsheetDocument(string docpath)
        {
            SpreadsheetDocument ExcelDoc = null;
            try
            {
                ExcelDoc = SpreadsheetDocument.Open(docpath, true);
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
                    ExcelDoc = SpreadsheetDocument.Open(docpath, true);
                }
            }

            return ExcelDoc;
        }





        private static Xdr.TwoCellAnchor GenerateDrawingPart( WorksheetPart worksheet, OpenXmlElement fromMarker, OpenXmlElement toMarker, string imageId,  string shapeId, string shapeName) {

            worksheet.AddNewPart<DrawingsPart>(imageId);        // changed imageId
            Xdr.TwoCellAnchor twoCellAnchor = new Xdr.TwoCellAnchor() { EditAs = Xdr.EditAsValues.OneCell };
            twoCellAnchor.Append(fromMarker);
            twoCellAnchor.Append(toMarker);
            Xdr.Picture picture = new Xdr.Picture();

            Xdr.NonVisualPictureProperties nonVisualPictureProperties = new Xdr.NonVisualPictureProperties();
            Xdr.NonVisualDrawingProperties nonVisualDrawingProperties = new Xdr.NonVisualDrawingProperties() { Id = (UInt32Value)Convert.ToUInt32(shapeId) , Name = shapeName };  // changed shapeId, shapeName

            Xdr.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties = new Xdr.NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks = new A.PictureLocks() { NoChangeAspect = true, NoChangeArrowheads = true };

            nonVisualPictureDrawingProperties.Append(pictureLocks);

            nonVisualPictureProperties.Append(nonVisualDrawingProperties);
            nonVisualPictureProperties.Append(nonVisualPictureDrawingProperties);

            Xdr.BlipFill blipFill = new Xdr.BlipFill();

            A.Blip blip = new A.Blip() { Embed = imageId };   // changed imageId

            A.BlipExtensionList blipExtensionList = new A.BlipExtensionList();

            A.BlipExtension blipExtension = new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };

            A14.UseLocalDpi useLocalDpi = new A14.UseLocalDpi() { Val = false };
            useLocalDpi.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            blipExtension.Append(useLocalDpi);

            blipExtensionList.Append(blipExtension);

            blip.Append(blipExtensionList);

            A.Stretch stretch = new A.Stretch();
            A.FillRectangle fillRectangle = new A.FillRectangle();

            stretch.Append(fillRectangle);

            blipFill.Append(blip);
            blipFill.Append(stretch);

            Xdr.ShapeProperties shapeProperties = new Xdr.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

            A.Transform2D transform2D = new A.Transform2D();
            A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents1 = new A.Extents() { Cx = 0L, Cy = 0L };

            transform2D.Append(offset1);
            transform2D.Append(extents1);
            A.PresetGeometry presetGeometry = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.NoFill noFill = new A.NoFill();

            A.ShapePropertiesExtensionList shapePropertiesExtensionList = new A.ShapePropertiesExtensionList();

            A.ShapePropertiesExtension shapePropertiesExtension = new A.ShapePropertiesExtension() { Uri = "{909E8E84-426E-40DD-AFC4-6F175D3DCCD1}" };

            A14.HiddenFillProperties hiddenFillProperties = new A14.HiddenFillProperties();
            hiddenFillProperties.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            A.SolidFill solidFill = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex = new A.RgbColorModelHex() { Val = "FFFFFF" };

            solidFill.Append(rgbColorModelHex);

            hiddenFillProperties.Append(solidFill);

            shapePropertiesExtension.Append(hiddenFillProperties);

            shapePropertiesExtensionList.Append(shapePropertiesExtension);

            shapeProperties.Append(transform2D);
            shapeProperties.Append(presetGeometry);
            shapeProperties.Append(noFill);
            shapeProperties.Append(shapePropertiesExtensionList);

            picture.Append(nonVisualPictureProperties);
            picture.Append(blipFill);
            picture.Append(shapeProperties);
            Xdr.ClientData clientData = new Xdr.ClientData();


            twoCellAnchor.Append(picture);
            twoCellAnchor.Append(clientData);
            return twoCellAnchor;
        }

        #endregion


    }
}

public class FilePossiblyTooBiGException : Exception
{
    public FilePossiblyTooBiGException(string message)
       : base(string.Format("ERROR was thrown because the system ran out of memory, a possible cause for this error could be that the file being proccessed is too large, the following error was thrown {0}", message))
    {
    }
}
