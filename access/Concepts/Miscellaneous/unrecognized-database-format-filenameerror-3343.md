---
title: Unrecognized database format <filename>. (Error 3343)
keywords: jeterr40.chm5003343
f1_keywords:
- jeterr40.chm5003343
ms.assetid: d917be92-c946-1764-9409-9368d011390a
ms.date: 06/08/2017
ms.localizationpriority: high
---


# Unrecognized database format \<filename\>. (Error 3343)

  

**Applies to:** Access 2013 | Access 2016

Possible causes:



- The specified file name is not a Microsoft Access database engine database.
    
- The specified file name is a device name, for example, a printer or a console.
    
- The database file has invalid header information or an unknown sort order.
    
- A commit is pending from another user but the lock file cannot be found.
    
- During a commit, you are attempting to write a Long value larger than the 2K maximum page size.
    
- The database is damaged. Compact the database and then try opening it again.
    

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
