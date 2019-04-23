---
title: OlDefaultFolders enumeration (Outlook)
keywords: vbaol11.chm3055
f1_keywords:
- vbaol11.chm3055
ms.prod: outlook
api_name:
- Outlook.OlDefaultFolders
ms.assetid: 1a17abd8-09b9-d3e1-2d93-0a4d5580a950
ms.date: 06/08/2017
localization_priority: Normal
---


# OlDefaultFolders enumeration (Outlook)

Specifies the folder type for a specified folder.



|Name|Value|Description|
|:-----|:-----|:-----|
| **olFolderCalendar**|9|The Calendar folder.|
| **olFolderConflicts**|19|The Conflicts folder (subfolder of the Sync Issues folder). Only available for an Exchange account.|
| **olFolderContacts**|10|The Contacts folder.|
| **olFolderDeletedItems**|3|The Deleted Items folder.|
| **olFolderDrafts**|16|The Drafts folder.|
| **olFolderInbox**|6|The Inbox folder.|
| **olFolderJournal**|11|The Journal folder.|
| **olFolderJunk**|23|The Junk E-Mail folder.|
| **olFolderLocalFailures**|21|The Local Failures folder (subfolder of the Sync Issues folder). Only available for an Exchange account.|
| **olFolderManagedEmail**|29|The top-level folder in the Managed Folders group. For more information on Managed Folders, see the Help in Microsoft Outlook. Only available for an Exchange account.|
| **olFolderNotes**|12|The Notes folder.|
| **olFolderOutbox**|4|The Outbox folder.|
| **olFolderSentMail**|5|The Sent Mail folder.|
| **olFolderServerFailures**|22|The Server Failures folder (subfolder of the Sync Issues folder). Only available for an Exchange account.|
| **olFolderSuggestedContacts**|30|The Suggested Contacts folder.|
| **olFolderSyncIssues**|20|The Sync Issues folder. Only available for an Exchange account.|
| **olFolderTasks**|13|The Tasks folder.|
| **olFolderToDo**|28|The To Do folder.|
| **olPublicFoldersAllPublicFolders**|18|The All Public Folders folder in the Exchange Public Folders store. Only available for an Exchange account.|
| **olFolderRssFeeds**|25|The RSS Feeds folder.|

## Remarks

Used as a parameter to the [NameSpace.GetSharedDefaultFolder](Outlook.NameSpace.GetSharedDefaultFolder.md), [NameSpace.GetDefaultFolder](Outlook.NameSpace.GetDefaultFolder.md), [Store.GetDefaultFolder](Outlook.Store.GetDefaultFolder.md), and [Folder.Add](Outlook.Folders.Add.md) methods. Also used by the [SharingItem.RequestFolder](Outlook.SharingItem.RequestedFolder.md) property.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
