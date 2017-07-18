using Microsoft.Office.Interop.Excel;
using SC.API.ComInterop;
using SC.API.ComInterop.Models;
using System;
using System.Configuration;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;

// Uses workbook made from SCExporter to upload data back to sharpcloud
namespace SCUploader
{
    class Program
    {
        static void Main(string[] args)
        {
            // Get info from config file
            var teamstoryid = ConfigurationManager.AppSettings["teamstoryid"];
            var portfolioid = ConfigurationManager.AppSettings["portfolioid"];
            var templateid = ConfigurationManager.AppSettings["templatIeid"];
            var userid = ConfigurationManager.AppSettings["user"];
            var passwd = ConfigurationManager.AppSettings["pass"];
            var URL = ConfigurationManager.AppSettings["URL"];
            var storyID = ConfigurationManager.AppSettings["story"];

            // Login and get story data from Sharpcloud
            var sc = new SharpCloudApi(userid, passwd, URL);
            var story = sc.LoadStory(storyID);
            // Load data from excel
            var app = new Application();
            Workbook sharpBook = app.Workbooks.Open
                (System.IO.Directory.GetParent
                (System.IO.Directory.GetParent
                (Environment.CurrentDirectory).ToString()).ToString() + "\\combine.xlsx");
            Worksheet item = sharpBook.Sheets[1];
            Worksheet relationship = sharpBook.Sheets[2];
            Range itemRange = item.UsedRange;
            int itemRow = itemRange.Rows.Count;
            int itemCol = itemRange.Columns.Count;
            Range relRange = relationship.UsedRange;
            int relRow = relRange.Rows.Count;
            int relCol = relRange.Columns.Count;

            // Parse through sheet 1
            for (int i = 2; i <= itemRow; i++)
            {
                // uploads item
                if (itemRange.Cells[i, 3] != null)
                {
                    // check to see if category is in the story
                    if (story.Category_FindByName(itemRange.Cells[i, 3].Value2) == null)
                    {
                        // adds new category to story if category is new
                        story.Category_AddNew(itemRange.Cells[i, 3].Value2);
                        // Sets color of category based on argb value
                        string[] colors = itemRange.Cells[i, itemCol - 1].Value2.ToString().Split('|');
                        var catColor = Color.FromArgb(Int32.Parse(colors[0]), Int32.Parse(colors[1]), Int32.Parse(colors[2]), Int32.Parse(colors[3]));
                        story.Category_FindByName(itemRange.Cells[i, 3].Value2).Color = catColor;
                    }
                    // check to see if item is in the story
                    if (story.Item_FindByName(itemRange.Cells[i,1].Value2) == null)
                    {
                        var catID = story.Category_FindByName(itemRange.Cells[i, 3].Value2);
                        // creates new item if item is not found
                        Item scItem = story.Item_AddNew(itemRange.Cells[i, 1].Value2);
                        // sets categiry for item
                        scItem.Category = catID; 
                        // checks subcategory of item
                        if(itemRange.Cells[i, itemCol - 3].Value2 != "null")
                        {
                            // checks to see if subcategory is in the story
                            if (catID.SubCategory_FindByName(itemRange.Cells[i, itemCol - 3].Value2) == null)
                            {
                                // adds subcategory to category of item if not found
                                catID.SubCategory_AddNew(itemRange.Cells[i, itemCol - 3].Value2);
                            }
                            // sets subcategory to the item
                            scItem.SubCategory = catID.SubCategory_FindByName(itemRange.Cells[i, itemCol - 3].Value2);
                            
                        }
                        // check to see if image path
                        if (itemRange.Cells[i, itemCol].Value2.ToString() != "null")
                        {
                            // uploads image to sharpcloud if image path found
                            FileInfo fileInfo = new FileInfo(itemRange.Cells[i, itemCol].Value2);
                            byte[] data = new byte[fileInfo.Length];
                            using (FileStream fs = fileInfo.OpenRead())
                            {
                                fs.Read(data, 0, data.Length);
                            }
                            scItem.ImageId = sc.UploadImageData(data, "", false);
                        }

                        /*string[] resources = itemRange.Cells[i, itemCol - 5].Value2.ToString().Split('|');
                        for(var z = 0; z < resources.Length - 1; z++)
                        {
                            string[] resLine = resources[i].Split('~');
                            string[] downLine = resLine[1].Split('*');
                            if(downLine.Length > 0)
                            {
                                scItem.Resource_AddFile(itemRange.Cells[i, itemCol].Value2.ToString() + downLine[0] + downLine[1], resLine[0], null);
                            }
                            else
                            {
                                scItem.Resource_AddName(resLine[0], null, resLine[1]);
                            }
                            
                        }

                        string[] tags = itemRange.Cells[i, itemCol - 4].Value2.ToString().Split('|');
                        for(var x = 0; x < tags.Length - 1; x++)
                        {
                            scItem.Tag_AddNew(tags[i]);
                        }
                        */
                    }
                   
                   
                }
            }
            // Parse through sheet 2
            for (int j = 2; j < relRow; j++)
            {
                // Establish relationships between 2 items
                var currentItem = story.Item_FindByName(relRange.Cells[j, 1].Value2.ToString());
                var nextItem = story.Item_FindByName(relRange.Cells[j, 2].Value2.ToString());
                var rel = story.Relationship_AddNew(currentItem, nextItem);
                rel.Comment = "";
                rel.Direction = Relationship.RelationshipDirection.None;
            }
            // Save story
            story.Save();
            //close and release excel
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.ReleaseComObject(itemRange);
            Marshal.ReleaseComObject(item);
            Marshal.ReleaseComObject(relationship);
            Marshal.ReleaseComObject(relRange);
            sharpBook.Close();
        }
    }
}
