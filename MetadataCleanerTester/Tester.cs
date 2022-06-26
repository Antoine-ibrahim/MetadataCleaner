using System;
using static MetadataCleaner.WordCleaner;
using static MetadataCleaner.PowerPointCleaner;
using static MetadataCleaner.ExcelCleaner;
using static MetadataCleaner.DocumentCleaner;

namespace MetadataCleanerTester
{
    class Tester
    {
        static void Main(string[] args)
        {
            // The tester is being used to clean Word documents first.
             string docPath = @"D:\Downloads\checkbox with bookmarks.docx";



            // MetadataCleaner.Excel.RemoveSystemProperties(docPath);      //done
            // MetadataCleaner.Excel.RemoveBuiltInProperties(docPath);     //done
            // MetadataCleaner.Excel.RemoveCustomProperties(docPath);                //done
            //MetadataCleaner.Excel.RemoveTrackedChanges(docPath);                 //DONE NEW
            //MetadataCleaner.Excel.RemoveCommentsDoc(docPath);                        //done
            //doReplaceEmbedsWithPictures(ExcelDoc);             
            //MetadataCleaner.Excel.RemoveHyperlinks(docPath);                      //done
            //MetadataCleaner.Excel.RemoveDefinedNames(docPath);                  // done 
            //MetadataCleaner.Excel.RemovePersonalInfo(docPath);           //done                 
            //MetadataCleaner.Excel.RemoveScinarios(docPath);                        //done 
            //  MetadataCleaner.Excel.RemoveHiddenRows(docPath);                       //done  NEED TO SEE WHAT TO DO IF THERE ARE FILTERS , MAY NEED TO BLCOK THE REMOVAL OF HIDDEN COLUMNS IF THERE ARE FILTERS 
            //MetadataCleaner.Excel.RemoveHiddenColumns(docPath);                    //Done NEW      // makes them visible instead of hidden
            ////doRemoveHiddenWorksheets(ExcelDoc);                               // seems hard save for later
            //MetadataCleaner.Excel.RemoveHiddenObjects(docPath);                    //done  
            ////   doRemoveLinksToExternalFiles(ExcelDoc);           
            //MetadataCleaner.Excel.RemoveCustomViews(docPath);                    //done               
            //MetadataCleaner.Excel.RemoveAutoFilter(docPath);                     // done , if there are hidden rows , they will be revealed using this function 
            // MetadataCleaner.Excel.RemoveCustomStyle(docPath);                     // DONE NEW  
            //MetadataCleaner.Excel.RemoveSparkLines(docPath);                      //done  
            //MetadataCleaner.Excel.RemoveHeaderAndFooter(docPath);                  //done
            //MetadataCleaner.Excel.RemoveSlicer(docPath);                           //done
            //MetadataCleaner.Word.Clean(docPath);
            // MetadataCleaner.Word.FixInvalidUri(docPath);
            // MetadataCleaner.Powerpoint.Clean(docPath);
            // MetadataCleaner.Excel.ReplaceEmbedsWithPictures(docPath);

            // MetadataCleaner.Excel.UnhideHiddenRows(docPath);



            //MetadataCleaner.Powerpoint.Clean(docPath);
            //MetadataCleaner.Excel.RemoveHyperlinks(docPath); ;
            // MetadataCleaner.DocumentCleaner.Clean(docPath);
            MetadataCleaner.WordCleaner.UnCheckCheckboxes(docPath);
            Console.WriteLine("Finished");
            Console.ReadLine();
        }

    }
}