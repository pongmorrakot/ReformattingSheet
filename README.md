# ReformattingSheet
Application to replace Reformatting Sheet used by UWDCC photographers
Unfinished Demo: https://drive.google.com/drive/folders/1244eHJNfOckgG59w18a7BJkoZrty0GXF?usp=sharing

Batch List - contains information(Collection Info, Date Created, Date Last Edited, Google Object ID) about all the Batches
     Sheet - Batch List - contains information about Batches
           - Deleted - contains information about Batches deleted by the user; batch in this sheet is being moved to folder Batch/RecycleBin; these batch will be actually deleted after a certain period of time.
        
Batches:

  -> Batch - each Batch has a unique name correspond to how it was named by user; First sheet of a Batch will always be Issue List and the rest of the sheet will be Issues listed in Issue List
        Sheet - Issue List - contains checklist for Reformatting QC check
              - Issues - contains checklist and information about each page in the issue
              
  -> Recycle Bin: contains Batches that were deleted by the user;
        -> Batch

Finished script:

	- update Upload Form to match entries BatchList

	- update "Last Edit" date in BatchList to match date where each Batch spreadsheet was last edited
	
Essentials function to implement:

	- take in user input from Google Form Upload Info to generate Batch spreadsheet and add entry to BatchList

Possible addition

	- manage Batch deletions; entry in BatchList is added/removed in relation to its corresponding Batch spreadsheet
	
	- button or sidebar that user can use to check box/boxes