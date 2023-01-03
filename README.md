# OutlookCollection
This is a small collection of some little helpers for Microsoft Outlook

* DupeKiller.vb

  Deletes duplicate mails within a folder structure. It does not matter where the duplicate mail is located, the second mail will be deleted. The duplicate is identified by: 
  
  * notes : the subject
  
  * apointments: the subject, the start date and the end date
  
  * emails: send date, the senders address and the subject
  
* MoveByTable.vb

* RemoveCopy.vb

  Removes the leading "Copy: " from the subject of an appointment or fills an empty subject ("") with a space (" "). This was very helpful when migrating and merging multiple calendars.
  
* SaveAttachments.vb 
  
  stores all attachments of a mailbox in a file system. It follows the same folder structure as the mailbox. Each e-mail has its own folder. The folder is named after the date of receipt in the format YYYYMMDD and the subject. The length of the folder is limited to 70 characters and is shortened if necessary.
