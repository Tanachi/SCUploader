using SC.API.ComInterop;
using SC.API.ComInterop.Models;
using System;
using System.Configuration;
using System.Drawing;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using System.Threading;
// Uses workbook made from SCExporter to upload data back to sharpcloud
namespace SCUploader
{
    class Program
    {
        public class Sharp
        {
            public SharpCloudApi sc { get; set; }
            public Story story { get; set; }
            public ExcelWorksheet sheet { get; set; }
            public int order { get; set; }
        }
        static Thread firstHalf, secondHalf;
        static void Main(string[] args)
        {
            string fileName = System.IO.Directory.GetParent(System.IO.Directory.GetParent(Environment.CurrentDirectory).ToString()).ToString() + "\\combine.xlsx";
            // Get info from config file
            var teamstoryid = ConfigurationManager.AppSettings["teamstoryid"];
            var portfolioid = ConfigurationManager.AppSettings["portfolioid"];
            var templateid = ConfigurationManager.AppSettings["templatIeid"];
            var userid = ConfigurationManager.AppSettings["user"];
            var passwd = ConfigurationManager.AppSettings["pass"];
            var URL = ConfigurationManager.AppSettings["URL"];
            var storyID = ConfigurationManager.AppSettings["story"];
            var storyTwo = ConfigurationManager.AppSettings["storyTwo"];

            // Login and get story data from Sharpcloud
            var sc = new SharpCloudApi(userid, passwd, URL);
            var story = sc.LoadStory(storyID);
            // Load data from excel
            using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo(fileName)))
            {
                // Generate data from worksheets
                var itemSheet = xlPackage.Workbook.Worksheets.First();
                var relSheet = xlPackage.Workbook.Worksheets.ElementAt(1);
                var relRows = relSheet.Dimension.End.Row;
                var relCols = relSheet.Dimension.End.Column;
                var itemRows = itemSheet.Dimension.End.Row;
                var itemColumns = itemSheet.Dimension.End.Column;
                // Creating the objects to be used in the threads
                object a, b;
                Sharp first = new Sharp();
                Sharp second = new Sharp();
                first.sheet = itemSheet; first.sc = sc; first.story = story; first.order = 0;
                second.sheet = itemSheet; second.sc = sc; second.story = story; second.order = 1;
                // Creating the threads
                firstHalf = new Thread(new ParameterizedThreadStart(addItems));
                secondHalf = new Thread(new ParameterizedThreadStart(addItems));
                firstHalf.Name = "firstHalf";
                secondHalf.Name = "secondHalf";
                // add attribute to story
                for (var k = 13; k < itemColumns; k++)
                {
                    string[] attribute = itemSheet.Cells[1, k].Value.ToString().Split('|');
                    if (story.Attribute_FindByName(attribute[0]) == null)
                    {
                        switch (attribute[1])
                        {
                            case "Text":
                                story.Attribute_Add(attribute[0], SC.API.ComInterop.Models.Attribute.AttributeType.Text);
                                break;
                            case "Numeric":
                                story.Attribute_Add(attribute[0], SC.API.ComInterop.Models.Attribute.AttributeType.Numeric);
                                break;
                            case "Date":
                                story.Attribute_Add(attribute[0], SC.API.ComInterop.Models.Attribute.AttributeType.Date);
                                break;
                            case "List":
                                story.Attribute_Add(attribute[0], SC.API.ComInterop.Models.Attribute.AttributeType.List);
                                break;
                            case "Location":
                                story.Attribute_Add(attribute[0], SC.API.ComInterop.Models.Attribute.AttributeType.Location);
                                break;
                        }
                        story.Attribute_FindByName(attribute[0]).Description = attribute[2];
                    }

                }
                // go through item sheet
                a = first;
                b = second;
                firstHalf.Start(a);
                secondHalf.Start(b);
                // Waits for all the items to be in the story before adding relationships
                firstHalf.Join();
                secondHalf.Join();
                Console.WriteLine("Adding Relationships");
                string[] attSplit = null;
                // Add attribute relationship to story
                for (var i = 6; i < relCols; i++)
                {
                    attSplit = relSheet.Cells[1, i].Value.ToString().Split('|');
                    if (story.RelationshipAttribute_FindByName(attSplit[0]) == null)
                    {
                        switch (attSplit[1])
                        {
                            case "Numeric":
                                story.RelationshipAttribute_Add(attSplit[0], RelationshipAttribute.RelationshipAttributeType.Numeric);
                                break;
                            case "Date":
                                story.RelationshipAttribute_Add(attSplit[0], RelationshipAttribute.RelationshipAttributeType.Date);
                                break;
                            case "List":
                                story.RelationshipAttribute_Add(attSplit[0], RelationshipAttribute.RelationshipAttributeType.List);
                                break;
                            case "Text":
                                story.RelationshipAttribute_Add(attSplit[0], RelationshipAttribute.RelationshipAttributeType.Text);
                                break;
                        }
                    }
                }
                // Add relationships to items
                for (int rowNum = 2; rowNum <= relRows; rowNum++)
                {
                    // Establish relationships between 2 items
                    var currentItem = story.Item_FindByName(relSheet.Cells[rowNum, 1].Value.ToString());
                    var nextItem = story.Item_FindByName(relSheet.Cells[rowNum, 2].Value.ToString());
                    var rel = story.Relationship_AddNew(currentItem, nextItem);
                    // Assign direction to relationship
                    switch (relSheet.Cells[rowNum, 3].Value.ToString())
                    {
                        case "None":
                            rel.Direction = Relationship.RelationshipDirection.None;
                            break;
                        case "AtoB":
                            rel.Direction = Relationship.RelationshipDirection.AtoB;
                            break;
                        case "BtoA":
                            rel.Direction = Relationship.RelationshipDirection.BtoA;
                            break;
                        case "Both":
                            rel.Direction = Relationship.RelationshipDirection.Both;
                            break;
                    }
                    string[] tagArray = null;
                    if (relSheet.Cells[rowNum, 4].Value != null && relSheet.Cells[rowNum,4].Value != "")
                    {
                        tagArray = relSheet.Cells[rowNum, 4].ToString().Split('|');
                        for(var i = 0; i < tagArray.Length - 1 ; i++)
                        {
                            rel.Tag_AddNew(tagArray[i]);

                        }
                    }
                    var relCom = story.Item_FindByName(relSheet.Cells[rowNum, 5].Value.ToString());
                    for(var k = 6; k < relCols; k++)
                    {
                        attSplit = relSheet.Cells[1, k].Value.ToString().Split('|');
                        rel.SetAttributeValue(story.RelationshipAttribute_FindByName(attSplit[0]), relSheet.Cells[rowNum,k].Value.ToString());
                    }
                }
            }
            story.Save();
        }
        public static void addItems(object Sharp)
        {
            Sharp sharp = Sharp as Sharp;
            var story = sharp.story;
            var sc = sharp.sc;
            var itemSheet = sharp.sheet;
            var itemRows = itemSheet.Dimension.End.Row;
            var itemColumns = itemSheet.Dimension.End.Column;
            int start = 2;
            if (Thread.CurrentThread.Name == "secondHalf")
            {
                start = 3;
            }
            for (int rowNum = start; rowNum <= itemRows; rowNum += 2) //selet starting row here
            {
                // check to see if category is in the story
                if (story.Category_FindByName(itemSheet.Cells[rowNum, 3].Value.ToString()) == null)
                {
                    // adds new category to story if category is new
                    story.Category_AddNew(itemSheet.Cells[rowNum, 3].Value.ToString());
                    // Sets color of category based on argb value
                    string[] colors = itemSheet.Cells[rowNum, 11].Value.ToString().Split('|');
                    var catColor = Color.FromArgb(Int32.Parse(colors[0]), Int32.Parse(colors[1]), Int32.Parse(colors[2]), Int32.Parse(colors[3]));
                    story.Category_FindByName(itemSheet.Cells[rowNum, 3].Value.ToString()).Color = catColor;

                }
                // Check to see if item is already in the story
                if (story.Item_FindByName(itemSheet.Cells[rowNum, 1].Value.ToString()) == null)
                {
                    var catID = story.Category_FindByName(itemSheet.Cells[rowNum, 3].Value.ToString());
                    Item scItem = story.Item_AddNew(itemSheet.Cells[rowNum, 1].Value.ToString());
                    scItem.Category = catID;
                    scItem.Description = itemSheet.Cells[rowNum, 2].GetValue<string>();
                    scItem.StartDate = Convert.ToDateTime(itemSheet.Cells[rowNum, 4].Value.ToString());
                    scItem.DurationInDays = itemSheet.Cells[rowNum, 2].GetValue<int>();
                    // checks subcategory of item
                    if (itemSheet.Cells[rowNum, 9].GetValue<string>() != "null")
                    {
                        // checks to see if subcategory is in the story
                        if (catID.SubCategory_FindByName(itemSheet.Cells[rowNum, 9].GetValue<string>()) == null)
                        {
                            // adds subcategory to category of item if not found
                            catID.SubCategory_AddNew(itemSheet.Cells[rowNum, 9].GetValue<string>());
                        }
                        // sets subcategory to the item
                        scItem.SubCategory = catID.SubCategory_FindByName(itemSheet.Cells[rowNum, 9].GetValue<string>());
                    }
                    // check to see if image path is there for item
                    if (itemSheet.Cells[rowNum, 12].GetValue<string>() != "null")
                    {
                        // uploads image to sharpcloud if image path found
                        if (File.Exists(itemSheet.Cells[rowNum, 12].GetValue<string>() + scItem.Name + ".jpg"))
                        {
                            FileInfo fileInfo = new FileInfo(itemSheet.Cells[rowNum, 12].GetValue<string>() + scItem.Name + ".jpg");
                            byte[] data = new byte[fileInfo.Length];
                            using (FileStream fs = fileInfo.OpenRead())
                            {
                                fs.Read(data, 0, data.Length);
                            }
                            scItem.ImageId = sc.UploadImageData(data, "", false);
                        }
                    }
                    // Check to see if item has resources
                    if (itemSheet.Cells[rowNum, 6].GetValue<string>() != "null")
                    {
                        string[] resources = itemSheet.Cells[rowNum, 6].GetValue<string>().Split('|');
                        for (var z = 0; z < resources.Length - 1; z++)
                        {
                            string[] resLine = resources[z].Split('~');
                            string[] downLine = resLine[1].Split('*');
                            // uploads file if there is a file extension
                            if (downLine.Length > 1)
                            {
                                if (File.Exists(itemSheet.Cells[rowNum, 12].GetValue<string>() + downLine[0] + downLine[1]))
                                    scItem.Resource_AddFile(itemSheet.Cells[rowNum, 12].GetValue<string>() + downLine[0] + downLine[1], resLine[0], null);
                            }
                            // adds url to another site
                            else
                            {
                                scItem.Resource_AddName(resLine[0], null, resLine[1]);
                            }
                        }
                    }
                    // Add Tags to the item
                    if (itemSheet.Cells[rowNum, 7].GetValue<string>() != "null")
                    {
                        string[] tags = itemSheet.Cells[rowNum, 7].GetValue<string>().Split('|');
                        for (var x = 0; x < tags.Length - 1; x++)
                        {
                            scItem.Tag_AddNew(tags[x]);
                        }
                    }
                    // Adds Panels to the item
                    if (itemSheet.Cells[rowNum, 8].GetValue<string>() != "null")
                    {
                        string[] panLine = itemSheet.Cells[rowNum, 8].GetValue<string>().Split('|');
                        for (var t = 0; t < panLine.Length - 1; t++)
                        {
                            
                            string[] panData = panLine[t].Split('@');
                            if (scItem.Panel_FindByTitle(panData[0]) == null)
                            {
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
                    // add attribute to the item
                    for (var j = 13; j < itemColumns; j++)
                    {
                        if (itemSheet.Cells[rowNum, j].Value != null && itemSheet.Cells[rowNum, j].Value.ToString() != "")
                        {
                            string[] attribute = itemSheet.Cells[1, j].Value.ToString().Split('|');
                            if (attribute[1] == "Date")
                            {
                                scItem.SetAttributeValue(story.Attribute_FindByName(attribute[0]), DateTime.FromOADate(Double.Parse(itemSheet.Cells[rowNum, j].Value.ToString())));
                            }
                            else if(attribute[1] == "Numeric")
                                scItem.SetAttributeValue(story.Attribute_FindByName(attribute[0]), Double.Parse(itemSheet.Cells[rowNum, j].Value.ToString()));
                            else
                                scItem.SetAttributeValue(story.Attribute_FindByName(attribute[0]), itemSheet.Cells[rowNum, j].Value.ToString());

                        }
                    }
                }

            }

        }
    }
}