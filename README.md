# vbascriptsforexcel

This script can be used to help format excel spreadsheets copied over from Atlassian Confluence tables into Excel.

One of the issues with tables copied from Confluence is that the bulleted notes in a cell get copied over into separate rows instead of the same row. This messes with the new spreadsheet.

This script leverages VBA macros in order to help fix this formatting issue. 

This script also helps parse columns that have one or more comma separated group labels and then groups specified content under those group labels. 

For example, consider the table below

User Story		Tag	Elaboration	Customer User Story 	Traceability
User Story Summary 1	Tag 1, Tag 2, Tag 3	1) Story line 1		Ref 1, Ref 2, Ref 3
						1.1)  Story line 2	
						1.2) Story line 3	

If we are interested in grouping the Customer User Story column by the tags (i.e. the group label in this case) in the Tag column or by Traceibility references in the Traceability column, we can achieve that using these scripts.

Usage:

1. Open Microsoft Excel.  
2. Select Tools->Macro->Visual Basic Editor
3. Copy the contents of the script into the Visual Basic Editor.
4. There are 2 main methods 

initializeTagWorksheet
----------------------

Purpose:  

1. Used to populate an empty worksheet with group labels for later grouping in the second step. 
2. This method will create a column for each group label, with the group label as the header. This is to be populated later with content relevant to each group label.

Arguments:

1. column to parse - The column that has comma separate group labels OR tags.
2. output worksheet number - The index of the worksheet that should hold the output data. 


formatUserStories
-----------------

Purpose: 
1. Used to fix Confluence copy issues discussed above i.e. multiple rows created when bullets/multilines are present in Conflunce tables,  format User Stories or any other similar content under groups.
2. Note that the first column in the last Row must have the string "End" in it. 

Arguments:

1. column to parse - The column that has comma separate group labels OR tags.
2. old elaboration column - This is the column with the multiple row problem on copy from Confluence that needs to be fixed. See Excel template for more information on how the Elaboration can be messed up on a copy from Confluence.
3. new elaboration column - Once formatting issues are fixed and cells are merged correctly, they are written to the new elaboration column. Upon completion of execution, this script will ask the user when or not it is ok to delete the older elaboration column i.e. the one with the messed up rows.
4. output worksheet number - The index of the worksheet that should hold the output data - content grouped under group labels.. 
5. user story summary - This is typically the first column in the spreadsheet. See template for more information on what this should look like. While there are multiple rows present, only the first row is populated and is visible. The script uses this information to under when the next Confluence table row begins. 

