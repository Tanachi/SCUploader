# SCUploader
Used with the spreadsheet from SCExporter to recreate the visuals from a story.

Spreadsheet must be in the folder where Program.CS is in

Enter info in App.config
# How to install

Add References

Project -> Add References
System.Configuration
System.Drawing

COM

Microsoft Excel 16.0 Object Library

Install Packages

Tools -> Nuget Package Manager -> Package Manager Console
Install-Package Newtonsoft.Json
Install-Package SharpCloud.ClientAPI -Version 1.0.18

# Issues:
If the program crashes, a instance of excel will still be up. Must close in task manager.

Panel data is stored in HTML. Working on trying to bring it to spreadsheet.

Program Crashes if you have any of the spreadsheets created from this program open during runtime.
