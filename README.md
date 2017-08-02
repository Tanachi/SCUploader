# SCUploader
Used with the spreadsheet from SCExporter to recreate the visuals from a story.

Spreadsheet must be in the folder where Program.CS is in

### How to install from Visual Studio

Create a new C# console application.

In the project folder, replace program.cs and app.config with the ones from this repo.

Insert combine.xlsx into the project folder

Add References

Project -> Add References

System.Configuration

System.Drawing

Install Packages

Tools -> Nuget Package Manager -> Package Manager Console 

Enter these lines in the console in this order.

Install-Package Newtonsoft.Json

Install-Package SharpCloud.ClientAPI

Install-Package EPPlus

Example sharpcloud Url

https://my.sharpcloud.com/html/#/story/Copy this Area/view/

Enter your Sharpcloud username, password, and story-id in the app.config file.

### Issues: 
Program Crashes if you have any of the spreadsheets created from this program open during runtime.

After uploading your spreadsheet to a sharpcloud story, the subcategories will appear twice. Refreshing this will remove this problem. 
