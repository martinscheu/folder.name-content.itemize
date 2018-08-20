# folder.name-content.itemize
Build a spreadsheet of folder and folder content names with a VBA macro
© various www resources and some code from myself

This Macro will read folder and files names, including all their subfolders and their file names	
Directories and File Names will be listed as direct links in a new worksheet.	
	
Config	
Parent Folder Path	Starting path for reading names.
	Note: Parent folder path will be removed from the names
	e.g. D:\Test or \\sites.inside-share.YourServer.com@ssl\davwwwroot\sites\Test
	
Target Worksheet Name	Name of the worksheet for the name tree. E.g. Input-Data
	
Include File Names	All file names in folders and subfolders will be read and listed
	
User Dialogue	
Yes: User needs to confirm worksheet name and will be asked if file names shall be read too
No: Macro will not confirm or ask for options and may override existing data
