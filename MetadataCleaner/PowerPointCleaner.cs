using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Xml.Linq;
using OpenXmlPowerTools;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.ExtendedProperties;
using System.Configuration;
using System.IO.Compression;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using A = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Presentation;

namespace MetadataCleaner
{
    public class PowerPointCleaner
    {
        #region Static Variables

        // Namespaces for various Word XML parts.
        //private static readonly XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        //private static readonly XNamespace x = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
        private static readonly XNamespace cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
        private static readonly XNamespace dc = "http://purl.org/dc/elements/1.1/";
        private static readonly XNamespace dcterms = "http://purl.org/dc/terms/";


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

            if (PRESENTATION_EXTENSIONS.Contains(extension))
            {
                using (PresentationDocument pptDoc = getPresentationDocument(pathToDocument))
                {
                    try
                    {
                        // Get XDocument to perform direct XML edit operations with.
                        XDocument coreXDoc = pptDoc.CoreFilePropertiesPart.GetXDocument();

                        doRemoveSystemProperties(pptDoc, coreXDoc);      //done 
                        doRemoveBuiltInProperties(pptDoc, coreXDoc);     //done
                        doRemoveCommentsDoc(pptDoc);                     //done
                        doReplaceEmbedsWithPictures(pptDoc);
                        doRemoveCustomProperties(pptDoc);                //done
                        doRemoveHiddenSlides(pptDoc);                    //done
                        doRemoveHiddenObjects(pptDoc);                   //done
                        doRemoveHyperlinks(pptDoc);                      //done
                        doRemoveSmallFont(pptDoc);                       //done 
                        doRemovePersonalInfo(pptDoc, coreXDoc);          //done           
                        doRemoveSpeakerNotes(pptDoc);                    //done
                        doRemoveOffsideobjects(pptDoc);                  //done
                        doRemoveAlternateText(pptDoc);                   //done
                        doRemoveHeaderAndFooter(pptDoc);                 //done

                        // Save the document XML back to their document parts.
                        coreXDoc.Save(pptDoc.CoreFilePropertiesPart.GetStream(FileMode.Create, FileAccess.Write));

                    }
                    catch (Exception ex)
                    {
                        if (ex.Message.Contains("System.OutOfMemoryException"))
                        {
                            throw new FilePossiblyTooBiGException(ex.Message);
                        }
                        if (ex.GetType().Name == "MissingFilePropertiesException")
                        {
                            throw new MissingFilePropertiesException("PowerPoint");
                        }
                        else
                        {
                            throw new Exception(ex.Message);
                        }
                    }

                    // TODO: Make sure all where statements are formatted correctly
                    // TODO: Add the comment header on the top of each function. 
                    // TODO: Make sure all global variables are starting with a capital letter, and local variables start with lowercase letter. 

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
            using (PresentationDocument pptDoc = getPresentationDocument(path))
            {
                try
                {
                    XDocument coreXDoc = pptDoc.CoreFilePropertiesPart.GetXDocument();

                    doRemoveSystemProperties(pptDoc, coreXDoc);

                    coreXDoc.Save(pptDoc.CoreFilePropertiesPart.GetStream(FileMode.Create, FileAccess.Write));
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


            using (PresentationDocument pptDoc = getPresentationDocument(path))
            {
                try
                {
                    XDocument coreXDoc = pptDoc.CoreFilePropertiesPart.GetXDocument();

                    doRemoveBuiltInProperties(pptDoc, coreXDoc);

                    coreXDoc.Save(pptDoc.CoreFilePropertiesPart.GetStream(FileMode.Create, FileAccess.Write));
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



        //==== RemoveComments ===================================================
        /// <summary>
        /// Removes all comments from the document and removes all references to the authors of the comments.
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================

        public static void RemoveComments(string path)
        {
            using (PresentationDocument pptDoc = getPresentationDocument(path))
            {
                try
                {
                    doRemoveCommentsDoc(pptDoc);
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




        //==== ReplaceEmbedsWithPictures ===================================================
        /// <summary>
        /// Replaces all embedded objects with a picture and removes all references to any embedded documents.  
        /// </summary>
        /// <param name="Path">The full path of the document that is being processed.</param>
        //=========================================================================
        public static void ReplaceEmbedsWithPictures(string path)
        {
            using (PresentationDocument pptDoc = getPresentationDocument(path))
            {
                try
                {
                    doReplaceEmbedsWithPictures(pptDoc);
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
        } // End ReplaceEmbedsWithPictures



        //==== RemoveCustomProperties ===================================================
        /// <summary>
        /// Removes all custom properties that were added to the document by users.   
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================

        public static void RemoveCustomProperties(string path)
        {
            using (PresentationDocument pptDoc = getPresentationDocument(path))
            {
                try
                {
                    doRemoveCustomProperties(pptDoc);
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




        //==== RemoveHiddenSlides ===================================================
        /// <summary>
        /// This removes all hidden slides in the powerpoint document.   
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================
        public static void RemoveHiddenSlides(string path)
        {
            using (PresentationDocument pptDoc = getPresentationDocument(path))
            {

                try
                {
                    doRemoveHiddenSlides(pptDoc);
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
        } // End RemoveHiddenSlides




        //==== RemoveHiddenObjects ===================================================
        /// <summary>
        /// This removes all hidden objects in the document, for example hidden pictures or text boxes.   
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================
        public static void RemoveHiddenObjects(string path)
        {
            using (PresentationDocument pptDoc = getPresentationDocument(path))
            {

                try
                {
                    doRemoveHiddenObjects(pptDoc);
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
        } // End RemoveHiddenObjects




        //==== RemoveHyperlinks ===================================================
        /// <summary>
        /// Remove any hyperlinks or on click links. This includes items such as a hyperlinked image.
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================
        public static void RemoveHyperlinks(string path)
        {
            using (PresentationDocument pptDoc = getPresentationDocument(path))
            {
                try
                {
                    doRemoveHyperlinks(pptDoc);
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



        //==== RemoveSmallFont ===================================================
        /// <summary>
        /// Remove any text that is invisible or unnoticable because it has a font size that is 1 or smaller.
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================
        public static void RemoveSmallFont(string path)
        {
            using (PresentationDocument pptDoc = getPresentationDocument(path))
            {
                try
                {
                    doRemoveSmallFont(pptDoc);
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
        } // End RemoveSmallFont



        //==== RemovePersonalInfo ===================================================
        /// <summary>
        /// Removes all personal info from the document and adds a setting that automatically removes personal information if the file is edited in the future.
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================

        public static void RemovePersonalInfo(string path)
        {
            using (PresentationDocument pptDoc = getPresentationDocument(path))
            {
                try
                {
                    XDocument coreXDoc = pptDoc.CoreFilePropertiesPart.GetXDocument();
                    doRemovePersonalInfo(pptDoc, coreXDoc);

                    coreXDoc.Save(pptDoc.CoreFilePropertiesPart.GetStream(FileMode.Create, FileAccess.Write));
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




        //==== RemoveSpeakerNotes ===================================================
        /// <summary>
        /// Removes the speaker notes from the bottom of each slide.
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================

        public static void RemoveSpeakerNotes(string path)
        {
            using (PresentationDocument pptDoc = getPresentationDocument(path))
            {
                try
                {
                    doRemoveSpeakerNotes(pptDoc);
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
        } // End RemoveSpeakerNotes





        //==== RemoveOffsideobjects ===================================================
        /// <summary>
        /// Removes Images and other objects that are not within the borders of the slide.
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================


        public static void RemoveOffsideobjects(string path)
        {
            using (PresentationDocument pptDoc = getPresentationDocument(path))
            {
                try
                {
                    doRemoveOffsideobjects(pptDoc);
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
        } // End RemoveOffsideobjects




        //==== RemoveAlternateText ===================================================
        /// <summary>
        /// Removes all title and description alternate text from the docuent.
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================

        public static void RemoveAlternateText(string path)
        {
            using (PresentationDocument pptDoc = getPresentationDocument(path))
            {
                try
                {
                    doRemoveAlternateText(pptDoc);
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
        } // End RemoveAlternateText




        //==== RemoveHeaderAndFooter ===================================================
        /// <summary>
        /// Removes all title and description alternate text from the docuent.
        /// </summary>
        /// <param name="path">The full path of the document that is being processed.</param>
        //=========================================================================

        public static void RemoveHeaderAndFooter(string path)
        {
            using (PresentationDocument pptDoc = getPresentationDocument(path))
            {
                try
                {
                    doRemoveHeaderAndFooter(pptDoc);
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
        } // End RemoveHeaderAndFooter



        #endregion

        #region Private Methods


        private static void doRemoveSystemProperties(PresentationDocument pptDoc, XDocument coreXDoc)
        {

            if (coreXDoc.Descendants().Count() <= 2 || pptDoc.ExtendedFilePropertiesPart.Properties.Count() <= 1)
            {
                throw new MissingFilePropertiesException("PowerPoint");
            }

            PresentationPart presentationPart = pptDoc.PresentationPart;
            ExtendedFilePropertiesPart docProperties = pptDoc.ExtendedFilePropertiesPart;

            // change the appversion to 12.0000
            OpenXmlElement AppVersion = docProperties.Properties
                .Where(a => a.LocalName == "AppVersion").FirstOrDefault();
            Ap.ApplicationVersion applicationVersion1 = new Ap.ApplicationVersion();
            applicationVersion1.Text = "16.0000";

            if (AppVersion != null)
            {
                AppVersion.InsertAfterSelf(applicationVersion1);
                AppVersion.Remove();
            }

            // edit the created and modified date to the date that the cleaner was run.
            XName created = dcterms + "created";
            XName modified = dcterms + "modified";

            coreXDoc.Descendants()
                .Where(x => x.Name == created || x.Name == modified)
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


        
        private static void doRemoveBuiltInProperties(PresentationDocument pptDoc, XDocument coreXDoc)
        {

            if (coreXDoc.Descendants().Count() <= 2 || pptDoc.ExtendedFilePropertiesPart.Properties.Count() <= 1)
            {
                throw new MissingFilePropertiesException("PowerPoint");
            }

            PresentationPart presentationPart = pptDoc.PresentationPart;
            ExtendedFilePropertiesPart ex = pptDoc.ExtendedFilePropertiesPart;

            OpenXmlElement company = ex.Properties
                .Where(a => a.LocalName == "Company").FirstOrDefault();

            OpenXmlElement hLinkBase = ex.Properties
                .Where(a => a.LocalName == "HyperlinkBase").FirstOrDefault();

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



        private static void doRemoveCommentsDoc(PresentationDocument pptDoc)
        {
            PresentationPart presentationPart = pptDoc.PresentationPart;
            // Set commentPart to the document WordprocessingCommentsPart, if it exists.
            List< SlidePart> slides = presentationPart.SlideParts.Where(a => a.SlideCommentsPart != null).ToList();
            foreach (SlidePart slide in slides)
            {
                slide.DeletePart(slide.SlideCommentsPart);
            }


            // Remove all personal information for Comment Authors and replace them wit a default "Author".
            CommentAuthorsPart CommentPart = pptDoc.PresentationPart.CommentAuthorsPart;
            if (CommentPart != null)
            {
                List<OpenXmlElement> authlist = CommentPart.CommentAuthorList.ToList();

                foreach (OpenXmlElement Author in authlist)
                {
                    UInt32Value id = Convert.ToUInt32(Author.GetAttribute("id", "").Value.ToString());
                    UInt32Value lastidx = Convert.ToUInt32(Author.GetAttribute("lastIdx", "").Value.ToString());
                    UInt32Value clrIdx = Convert.ToUInt32(Author.GetAttribute("clrIdx", "").Value.ToString());

                    CommentAuthor commentAuthor = new CommentAuthor() { Id = id, Name = "Author", Initials = "A", LastIndex = lastidx, ColorIndex = clrIdx };
                    if (Author != null)
                    {
                        Author.InsertAfterSelf(commentAuthor);
                        Author.Remove();
                    }
                }
            }

        } // End doRemoveCommentsDoc

        private static void doReplaceEmbedsWithPictures(PresentationDocument pptDoc)
        {
            PresentationPart presentationPart = pptDoc.PresentationPart;

            // get all slides with Embeds       
            List<SlidePart> SlidesWithEmbeds = presentationPart.SlideParts
                .Where(a => a.Slide.Descendants()
                .Where(b => b.GetType().Name.ToString() == "OleObject").Count() > 0).ToList();

            foreach (SlidePart slide in SlidesWithEmbeds)
            {
                // for each slide with Objects, get a list of Graphics frames ( containers of the objects )
                List<OpenXmlElement> graphicFrameList = slide.Slide.Descendants().Where(b => b.LocalName == "graphicFrame").ToList();

                foreach (OpenXmlElement graphicFrame in graphicFrameList)
                {
                    // get the id reference id of the ole object.
                    OpenXmlElement oleObject = graphicFrame.Descendants().Where(b => b.GetType().Name.ToString() == "OleObject").FirstOrDefault();
                    string id = oleObject.GetAttributes()
                        .Where(a => a.LocalName == "id").FirstOrDefault().Value;
                    // ReferenceRelationship hlr = slide.GetReferenceRelationship(id);
                    
                    // delete the referenced embededed object
                    //slide.DeletePart(id
                    try
                    {

                        ExternalRelationship ext = slide.GetExternalRelationship(id);
                        slide.DeleteExternalRelationship(ext);
                    }
                    catch {

                    }


                    OpenXmlElement picturePart = graphicFrame.Descendants().Where(b => b.GetType().Name.ToString() == "Picture").FirstOrDefault();
                    // replace the embed with a pic 

                    if (picturePart != null)
                    {
                            OpenXmlElement PicClone = picturePart.CloneNode(true);

                            OpenXmlElement drawingProps= PicClone.Descendants().Where(a => a.GetType().Name == "NonVisualPictureDrawingProperties").FirstOrDefault();

                        if (drawingProps != null)
                        {
                            OpenXmlElement existingPicLock = drawingProps.Descendants().Where(a => a.GetType().Name == "PictureLocks").FirstOrDefault();
                            if (existingPicLock != null)
                            {
                                existingPicLock.Remove();
                            }
                            A.PictureLocks pictureLocks = new A.PictureLocks() { NoChangeAspect = true };
                            drawingProps.Append(pictureLocks);
                        }
                            


                            graphicFrame.InsertAfterSelf(PicClone);
                    }
                  
                    graphicFrame.Remove();
                    List< VmlDrawingPart> vmls = slide.VmlDrawingParts.ToList();

                    for (int i =0; i < vmls.Count(); i++)
                    {
                        slide.DeletePart(vmls[i]);
                    }


                }

                // remove dirty attributes from slides elements
                List<OpenXmlElement> DirtyElements = slide.Slide.Descendants().Where(b => b.GetAttributes().Where(c => c.LocalName == "dirty").FirstOrDefault().Value == "0").ToList();
                foreach (OpenXmlElement element in DirtyElements)
                {
                    OpenXmlAttribute att = element.GetAttributes().Where(c => c.LocalName == "dirty").FirstOrDefault();
                    element.RemoveAttribute(att.LocalName, att.NamespaceUri);
                }
            }




        } // End doReplaceEmbedsWithPictures










        private static void doRemoveCustomProperties(PresentationDocument pptDoc)
        {
            // removes custom Properties
            CustomFilePropertiesPart CustomProperties = pptDoc.CustomFilePropertiesPart;
            pptDoc.DeletePart(CustomProperties);
        } // End doRemoveCustomProperties





        private static void doRemoveHiddenSlides(PresentationDocument pptDoc)
        {
            PresentationPart presentationPart = pptDoc.PresentationPart;
            List<OpenXmlElement> sldList = presentationPart.Presentation.Descendants().Where(a => a.LocalName == "sldId").ToList();  // gets the list of slide list from the presentation.xml document 

            foreach (SlideId slide in sldList)
            {
                string slideRelId = slide.RelationshipId;
                SlidePart slidePart = presentationPart.GetPartById(slideRelId) as SlidePart;

                if (slidePart.Slide.GetAttributes().Where(a => a.LocalName == "show").FirstOrDefault().Value == "0")
                {
                  //  Console.WriteLine("found a hidden slide");
                    presentationPart.DeletePart(slidePart);
                    slide.Remove();

                }
            }


        } // End doRemoveHiddenSlides





        private static void doRemoveHiddenObjects(PresentationDocument pptDoc)
        {

            PresentationPart presentationPart = pptDoc.PresentationPart;
            // get and remove all hyperlink elements 
            List<SlidePart> slideParts = presentationPart.SlideParts.ToList();
            foreach (SlidePart slidePart in slideParts)
            {
                List<OpenXmlElement> hiddenElements = slidePart.Slide.Descendants()
                    .Where(a=>a.LocalName.ToString() == "cNvPr"
                    && a.GetAttributes().Where(c => c.LocalName.ToString() == "hidden").FirstOrDefault().Value == "1").ToList();

                foreach (OpenXmlElement element in hiddenElements)
                {
                    OpenXmlElement hiddenElementContainer= element.Parent.Parent;
                    hiddenElementContainer.Remove();
                }

            }


        } // End doRemoveHiddenObjects

        private static void doRemoveHyperlinks(PresentationDocument pptDoc)
        {
            // Removes hyperlinks from document.
            PresentationPart presentationPart = pptDoc.PresentationPart;

            // get all slides with hyperlinks         
            List<SlidePart> SlidesWithHyperlinks = presentationPart.SlideParts
                .Where(a => a.Slide.Descendants()
                .Where(b => b.LocalName == "hlinkClick").Count() > 0).ToList();
                       
            foreach (SlidePart slide in SlidesWithHyperlinks)
            {
                // for each slide with hyperlinks , remove the hyperlink and references
                List<OpenXmlElement> hyperlinkList = slide.Slide.Descendants().Where(b => b.LocalName == "hlinkClick").ToList();
                foreach (OpenXmlElement hyperlink in hyperlinkList)
                {
                    string id = hyperlink.GetAttributes()
                        .Where(a => a.LocalName == "id").FirstOrDefault().Value;
                    if (!String.IsNullOrEmpty(id))
                    {
                        try
                        {
                            ReferenceRelationship hlr = slide.GetReferenceRelationship(id);
                            hyperlink.Remove();
                            slide.DeleteReferenceRelationship(hlr);
                        }
                        catch {
                            hyperlink.Remove();
                        }
                        

                    }
                }

                // remove dirty attributes from slides elements
                List<OpenXmlElement> DirtyElements = slide.Slide.Descendants().Where(b => b.GetAttributes().Where(c => c.LocalName == "dirty").FirstOrDefault().Value == "0").ToList();
                foreach (OpenXmlElement element in DirtyElements)
                {
                    OpenXmlAttribute att = element.GetAttributes().Where(c => c.LocalName == "dirty").FirstOrDefault();
                    element.RemoveAttribute(att.LocalName, att.NamespaceUri);
                }
            }
        } // End doRemoveHyperlinks

        private static void doRemoveSmallFont(PresentationDocument pptDoc)
        {
            // Removes text with small font (1pt or smaller).
            PresentationPart presentationPart = pptDoc.PresentationPart;

            // get and remove all hyperlink elements 
            List<SlidePart> allslides = presentationPart.SlideParts.ToList();
            foreach (SlidePart slide in allslides)
            {
                // removes all run elements with text smaller then 3pts
                List<OpenXmlElement> smallfont = slide.Slide.Descendants()
                    .Where(b => b.LocalName == "rPr"
                    && b.GetAttributes().Where(c => c.LocalName == "sz").FirstOrDefault().Value != null
                    && Convert.ToInt32(b.GetAttributes().Where(c => c.LocalName == "sz").FirstOrDefault().Value) <= 300).ToList();

                foreach (OpenXmlElement rpr in smallfont)
                {
                    string size = rpr.GetAttributes().Where(c => c.LocalName == "sz").FirstOrDefault().Value;
                    OpenXmlElement parent = rpr.Ancestors().Where(a => a.GetType().Name == "Run").FirstOrDefault();
                    parent.Remove();
                }

                // remove all dirty elements
                List<OpenXmlElement> DirtyElements = slide.Slide.Descendants().Where(b => b.GetAttributes().Where(c => c.LocalName == "dirty").FirstOrDefault().Value == "0").ToList();
                foreach (OpenXmlElement element in DirtyElements)
                {
                    OpenXmlAttribute att = element.GetAttributes().Where(c => c.LocalName == "dirty").FirstOrDefault();
                    element.RemoveAttribute(att.LocalName, att.NamespaceUri);
                }

            }
        } // End doRemoveSmallFont

        private static void doRemovePersonalInfo(PresentationDocument pptDoc, XDocument coreXDoc)
        {

            PresentationPart presentationPart = pptDoc.PresentationPart;
            ExtendedFilePropertiesPart ExtendedProperties = pptDoc.ExtendedFilePropertiesPart;

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

            // Remove all personal information for Comment Authors and replace them wit a default "Author".
            CommentAuthorsPart CommentPart = pptDoc.PresentationPart.CommentAuthorsPart;

            if (CommentPart != null) {
                List<OpenXmlElement> authlist = CommentPart.CommentAuthorList.ToList();

                foreach (OpenXmlElement Author in authlist)
                {
                    UInt32Value id = Convert.ToUInt32(Author.GetAttribute("id", "").Value.ToString());
                    UInt32Value lastidx = Convert.ToUInt32(Author.GetAttribute("lastIdx", "").Value.ToString());
                    UInt32Value clrIdx = Convert.ToUInt32(Author.GetAttribute("clrIdx", "").Value.ToString());

                    CommentAuthor commentAuthor = new CommentAuthor() { Id = id, Name = "Author", Initials = "A", LastIndex = lastidx, ColorIndex = clrIdx };
                    Author.InsertAfterSelf(commentAuthor);
                    Author.Remove();
                }

            }

            // Turns on the setting that removes personal info upon document save
            presentationPart.Presentation.RemovePersonalInfoOnSave= true;

        } // End doRemovePersonalInfo

        private static void doRemoveSpeakerNotes(PresentationDocument pptDoc)
        {
            // Removes text with small font (1pt or smaller).
            PresentationPart presentationPart = pptDoc.PresentationPart;
            // get and remove all hyperlink elements 
            List<SlidePart> slideParts = presentationPart.SlideParts.ToList();
            foreach (SlidePart slidePart in slideParts)
            {
                try
                {
                    // fined the part of the slide notes that has the notes placeholder
                    NotesSlidePart notes = slidePart.NotesSlidePart;
                    OpenXmlElement NotesPlaceholder = notes.NotesSlide.Descendants()
                        .Where(a => a.LocalName == "cNvPr"
                        && a.GetAttributes().Where(b => b.LocalName == "name").FirstOrDefault().Value.Contains("Notes Placeholder")).FirstOrDefault();

                    //Get the Text content of the slide notes and replace it with an empty textbody 
                    OpenXmlElement notesContainer = NotesPlaceholder.Ancestors()
                        .Where(a => a.GetType().Name.ToString() == "Shape").FirstOrDefault();

                    OpenXmlElement SlideNotes = notesContainer.Descendants()
                        .Where(a => a.GetType().Name == "TextBody").FirstOrDefault();

                    TextBody textBody = new TextBody();
                    A.BodyProperties bodyProperties = new A.BodyProperties();
                    A.ListStyle listStyle = new A.ListStyle();
                    A.Paragraph paragraph = new A.Paragraph();
                    A.EndParagraphRunProperties endParagraphRunProperties = new A.EndParagraphRunProperties() { Language = "en-US" };
                    paragraph.Append(endParagraphRunProperties);
                    textBody.Append(bodyProperties);
                    textBody.Append(listStyle);
                    textBody.Append(paragraph);

                    SlideNotes.InsertAfterSelf(textBody);
                    SlideNotes.Remove();
                    //Console.WriteLine("notes found on this slide");

                    List<OpenXmlElement> ShapeProperties = notes.NotesSlide.Descendants().Where(a => a.Parent.GetType().Name == "TransformGroup" && (a.GetType().Name == "ChildOffset" || a.GetType().Name == "ChildExtents")).ToList();
                    foreach (OpenXmlElement item in ShapeProperties)
                    {
                        item.Remove();
                    }
                }
                catch (Exception)
                {
                }
            }
        } // End doRemoveSpeakerNotes

        private static void doRemoveOffsideobjects(PresentationDocument pptDoc)
        {
            int maxXLength = 12192000;
            int maxYLength = 6858000;
            // Removes text with small font (1pt or smaller).
            PresentationPart presentationPart = pptDoc.PresentationPart;
            // get and remove all hyperlink elements 
            List<SlidePart> slideParts = presentationPart.SlideParts.ToList();

            foreach (SlidePart slidePart in slideParts)
            {
                List<OpenXmlElement> slideElementProperties = slidePart.Slide.Descendants().Where(a => a.GetType().Name == "ShapeProperties").ToList();
                foreach (OpenXmlElement slideElement in slideElementProperties)
                {
                    try
                    {
                       OpenXmlElement offSet= slideElement.Descendants().Where(a => a.GetType().Name == "Offset").FirstOrDefault();
                        string x = offSet.GetAttributes().Where(a => a.LocalName == "x").FirstOrDefault().Value;
                        string y = offSet.GetAttributes().Where(a => a.LocalName == "y").FirstOrDefault().Value;

                        OpenXmlElement extents=  slideElement.Descendants().Where(a => a.GetType().Name == "Extents").FirstOrDefault();
                        string cx = extents.GetAttributes().Where(a => a.LocalName == "cx").FirstOrDefault().Value;
                        string cy = extents.GetAttributes().Where(a => a.LocalName == "cy").FirstOrDefault().Value;


                        if (x != null && y!=null )
                        {
                            if (Convert.ToInt64(x) < 0 ||  Convert.ToInt64(y) < 0)
                            {
                             //   Console.WriteLine("this element is starting outside the bounds of the slide");
                                slideElement.Parent.Remove();
                            }
                           else if ((Convert.ToInt64(x) + Convert.ToInt64(cx) > maxXLength) || (Convert.ToInt64(y) + Convert.ToInt64(cy) > maxYLength))
                            {
                              //  Console.WriteLine("this elemeent ends outside the bounds of the slide");
                                slideElement.Parent.Remove();
                            }

                        }


                    }
                    catch {

                    }                    
                }

            }
            
        } // End doRemoveOffsideobjects





        private static void doRemoveAlternateText(PresentationDocument pptDoc)
        {
            PresentationPart presentationPart = pptDoc.PresentationPart;
            List<SlidePart> allslides = presentationPart.SlideParts.ToList();
            foreach (SlidePart slide in allslides)
            {
                // removes all footer elements 
                List<OpenXmlElement> itemsWithAltText = slide.Slide.Descendants()
                    .Where(b => b.LocalName == "cNvPr"
                    && b.GetAttributes().Where(c => c.LocalName == "descr").FirstOrDefault().Value != null
                    || (b.GetAttributes().Where(c => c.LocalName == "title").FirstOrDefault().Value !=null)).ToList();

                foreach (OpenXmlElement item in itemsWithAltText)
                {
                    OpenXmlAttribute desc = item.GetAttributes().Where(c => c.LocalName == "descr").FirstOrDefault();
                    OpenXmlAttribute title = item.GetAttributes().Where(c => c.LocalName == "title").FirstOrDefault();
                    if (!string.IsNullOrEmpty(desc.Value))
                    {
                        item.RemoveAttribute(desc.LocalName, desc.NamespaceUri);
                    }

                    if (!string.IsNullOrEmpty(title.Value))
                    {
                        item.RemoveAttribute(title.LocalName, title.NamespaceUri);
                    }

                }
            }



        } // End doRemoveAlternateText






        private static void doRemoveHeaderAndFooter(PresentationDocument pptDoc)
        {
            PresentationPart presentationPart = pptDoc.PresentationPart;

            // get a list of all the handoutmasters and notes master pages 
            List<OpenXmlElement> HandoutMasterList = presentationPart.Presentation.Descendants().Where(a => a.LocalName == "handoutMasterId").ToList();  
            List<OpenXmlElement> notesMasterlist = presentationPart.Presentation.Descendants().Where(a => a.LocalName == "notesMasterId").ToList(); 

            // delete the handoutmaster and notesmasters by id reference and from the presentation
            foreach (HandoutMasterId handout in HandoutMasterList)
            {
                string handoutId = handout.Id;// .RelationshipId;
                
                OpenXmlPart handoutmaster= presentationPart.GetPartById(handoutId);
                presentationPart.DeletePart(handoutmaster);
                handout.Remove();
            }

            foreach (NotesMasterId note in notesMasterlist)
            {
                string NoteId = note.Id;// .RelationshipId;

                OpenXmlPart noteMaster = presentationPart.GetPartById(NoteId);
                presentationPart.DeletePart(noteMaster);
                note.Remove();
            }

            // remove the containers for the notmaster and handout master lists
            OpenXmlElement notesMasterContainer = presentationPart.Presentation.Descendants().Where(a => a.LocalName == "notesMasterIdLst").FirstOrDefault();
            if (notesMasterContainer != null)
            {
                notesMasterContainer.Remove();
            }
            
            OpenXmlElement handoutMasterContainer = presentationPart.Presentation.Descendants().Where(a => a.LocalName == "handoutMasterIdLst").FirstOrDefault();
            if (handoutMasterContainer != null)
            {
                handoutMasterContainer.Remove();
            }

            // clean up the footer information from all the slides

            List<SlidePart> allslides = presentationPart.SlideParts.ToList();
            foreach (SlidePart slide in allslides)
            {
                // removes all footer elements 
                List<OpenXmlElement> FooterProperties = slide.Slide.Descendants()
                    .Where(b => b.LocalName == "cNvPr"
                    && b.GetAttributes().Where(c => c.LocalName == "name").FirstOrDefault().Value != null
                    && (b.GetAttributes().Where(c => c.LocalName == "name").FirstOrDefault().Value.Contains("Date Placeholder")
                    || b.GetAttributes().Where(c => c.LocalName == "name").FirstOrDefault().Value.Contains("Footer Placeholder")
                    || b.GetAttributes().Where(c => c.LocalName == "name").FirstOrDefault().Value.Contains("Slide Number Placeholder")) ).ToList();

                foreach (OpenXmlElement footerProperty in FooterProperties)
                {
                   OpenXmlElement shape= footerProperty.Parent.Parent;
                    shape.Remove();
                }
            }


            // clean up all the slide layouts and the slide master  of the footer information 

            List<OpenXmlElement> slideLayoutIds= allslides.FirstOrDefault().SlideLayoutPart.SlideMasterPart.SlideMaster.Descendants().Where(a => a.GetType().Name == "SlideLayoutId").ToList();
            foreach (SlideLayoutId slideLayout in slideLayoutIds)
            {
                string slideRelId = slideLayout.RelationshipId;

                SlideLayoutPart slidePart = allslides.FirstOrDefault().SlideLayoutPart.SlideMasterPart.GetPartById(slideRelId) as SlideLayoutPart;

                OpenXmlElement layoutFooterProps = slidePart.SlideLayout.Descendants()
                    .Where(a => a.LocalName == "cNvPr"
                     && a.GetAttributes().Where(c => c.LocalName == "name").FirstOrDefault().Value != null
                     && a.GetAttributes().Where(c => c.LocalName == "name").FirstOrDefault().Value.Contains("Footer Placeholder")).FirstOrDefault();

               

                if (layoutFooterProps != null) {

                    OpenXmlElement Layoutshape = layoutFooterProps.Ancestors().Where(a => a.GetType().Name == "Shape").FirstOrDefault();
                    OpenXmlElement LayoutfooterParagraph = Layoutshape.Descendants().Where(a => a.GetType().Name == "Paragraph").FirstOrDefault();
                    A.Paragraph para = new A.Paragraph();
                    A.EndParagraphRunProperties endParagraphRunProperties = new A.EndParagraphRunProperties() { Language = "en-US" };

                    para.Append(endParagraphRunProperties);
                    LayoutfooterParagraph.InsertAfterSelf(para);
                    LayoutfooterParagraph.Remove();
                }



                OpenXmlElement HeaderFooterElement = slidePart.SlideMasterPart.SlideMaster.ChildElements.Where(a => a.GetType().Name == "HeaderFooter").FirstOrDefault();
                if (HeaderFooterElement != null)
                {
                    HeaderFooterElement.Remove();
                }




                OpenXmlElement slideMasterFooterProps = slidePart.SlideMasterPart.SlideMaster.Descendants()
                     .Where(a => a.LocalName == "cNvPr"
                     && a.GetAttributes().Where(c => c.LocalName == "name").FirstOrDefault().Value != null
                     && a.GetAttributes().Where(c => c.LocalName == "name").FirstOrDefault().Value.Contains("Footer Placeholder")).FirstOrDefault();

                if (slideMasterFooterProps != null)
                {
                    OpenXmlElement Mastershape = slideMasterFooterProps.Ancestors().Where(a => a.GetType().Name == "Shape").FirstOrDefault();

                    OpenXmlElement MasterfooterParagraph = Mastershape.Descendants().Where(a => a.GetType().Name == "Paragraph").FirstOrDefault();
                    A.Paragraph paragraph = new A.Paragraph();
                    A.EndParagraphRunProperties endParagraphRunProperty = new A.EndParagraphRunProperties() { Language = "en-US" };
                    paragraph.Append(endParagraphRunProperty);
                    MasterfooterParagraph.InsertAfterSelf(paragraph);
                    MasterfooterParagraph.Remove();
                }
            }
        } // End doRemoveHeaderAndFooter


        #endregion


        #region Util Methods

        // area for helper functions 

        public static PresentationDocument getPresentationDocument(string docpath)
        {
            PresentationDocument pptDoc = null;
            try
            {
                pptDoc = PresentationDocument.Open(docpath, true);
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
                    pptDoc = PresentationDocument.Open(docpath, true);
                }
            }
           
             return pptDoc;
        }
        #endregion

    }
}
