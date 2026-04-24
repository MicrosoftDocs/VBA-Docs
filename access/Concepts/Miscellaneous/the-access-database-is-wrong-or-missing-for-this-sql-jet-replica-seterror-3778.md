---
title: The Access database is wrong or missing for this SQL/Jet replica set. (Error 3778)
keywords: jeterr40.chm5003778
f1_keywords:
- jeterr40.chm5003778
ms.assetid: 21c104f6-06cb-f3b5-aa9e-098dca92c223
ms.date: 06/08/2017
ms.localizationpriority: medium
---


# The Access database is wrong or missing for this SQL/Jet replica set. (Error 3778)

  

**Applies to:** Access 2013 | Access 2016

The Microsoft Access database specified is wrong or missing because:



- The data source path supplied in the link server on the SQL Server points to an invalid replica for this publication. The replica may have been replaced, damaged, or not created properly.
    
- The hub row is missing or is not valid.
    
- The database specified is not replicable.
    

The solution is to re-initialize your Jet Subscriber using the Re-Initialize tools on the SQL Server.

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]