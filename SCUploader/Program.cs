using Microsoft.Office.Interop.Excel;
using SC.API.ComInterop;
using SC.API.ComInterop.Models;
using System;
using System.Configuration;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;

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
                        string[] colors = itemRange.Cells[i, 10].Value2.ToString().Split('|');
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
                        scItem.Description = itemRange.Cells[i, 2].Value2;
                        // checks subcategory of item
                        if(itemRange.Cells[i, 8].Value2 != "null")
                        {
                            // checks to see if subcategory is in the story
                            if (catID.SubCategory_FindByName(itemRange.Cells[i, 8].Value2) == null)
                            {
                                // adds subcategory to category of item if not found
                                catID.SubCategory_AddNew(itemRange.Cells[i, 8].Value2);
                            }
                            // sets subcategory to the item
                            scItem.SubCategory = catID.SubCategory_FindByName(itemRange.Cells[i, 8].Value2);
                            
                        }
                        // check to see if image path
                        if (itemRange.Cells[i, 11].Value2.ToString() != "null")
                        {
                            // uploads image to sharpcloud if image path found
                            FileInfo fileInfo = new FileInfo(itemRange.Cells[i, 11].Value2 + scItem.Name + ".jpg");
                            byte[] data = new byte[fileInfo.Length];
                            using (FileStream fs = fileInfo.OpenRead())
                            {
                                fs.Read(data, 0, data.Length);
                            }
                            scItem.ImageId = sc.UploadImageData(data, "", false);
                        }
                        // Add resources to item
                        if (itemRange.Cells[i, 5].Value2.ToString() != "null")
                        {
                            string[] resources = itemRange.Cells[i, 5].Value2.Split('|');
                            for (var z = 0; z < resources.Length - 1; z++)
                            {
                                string[] resLine = resources[z].Split('~');
                                string[] downLine = resLine[1].Split('*');
                                // uploads file if there is a file extension
                                if (downLine.Length > 1)
                                {
                                    scItem.Resource_AddFile(itemRange.Cells[i, 11].Value2.ToString() + downLine[0] + downLine[1], resLine[0], null);
                                }
                                // adds url to another site
                                else
                                {
                                    scItem.Resource_AddName(resLine[0], null, resLine[1]);
                                }
                            }
                        }
                        // adds Tags to the item
                        if(itemRange.Cells[i,6].Value2.ToString() != "null")
                        {
                            string[] tags = itemRange.Cells[i, 6].Value2.Split('|');
                            for (var x = 0; x < tags.Length - 1; x++)
                            {
                                scItem.Tag_AddNew(tags[x]);
                            }
                        }
                        // Adds Panels to the item
                        if(itemRange.Cells[i,7].Value2.ToString() != "null")
                        {
                            string[] panLine = itemRange.Cells[i, 7].Value2.Split('|');
                            for(var t = 0; t< panLine.Length - 1; t++)
                            {
                                string[] panData = panLine[t].Split('@');
                                // sets panel type based off string
                                switch (panData[1])
                                {
                                    case "RichText":
                                        scItem.Panel_Add(panData[0], Panel.PanelType.RichText, panData[2]);
                                        break;
                                    case "Attribute":
                                        scItem.Panel_Add(panData[0], Panel.PanelType.Attribute, panData[2]);
                                        break;
                                    case "CustomResource":
                                        scItem.Panel_Add(panData[0], Panel.PanelType.CustomResource, panData[2]);
                                        break;
                                    case "HTML":
                                        scItem.Panel_Add(panData[0], Panel.PanelType.HTML, panData[2]);
                                        break;
                                    case "Image":
                                        scItem.Panel_Add(panData[0], Panel.PanelType.Image, panData[2]);
                                        break;
                                    case "Video":
                                        scItem.Panel_Add(panData[0], Panel.PanelType.Video, panData[2]);
                                        break;
                                    case "Undefined":
                                        scItem.Panel_Add(panData[0], Panel.PanelType.Undefined, panData[2]);
                                        break;
                                }
                            }
                        }
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
