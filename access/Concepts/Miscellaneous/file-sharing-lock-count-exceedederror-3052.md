---
title: File sharing lock count exceeded. (Error 3052)
keywords: jeterr40.chm5003052
f1_keywords:
- jeterr40.chm5003052
ms.assetid: 682df85c-6e2e-26d4-d035-d787de5672ae
ms.date: 06/08/2019
ms.localizationpriority: medium
---


# File sharing lock count exceeded. (Error 3052)

  

**Applies to:** Access 2013 | Access 2016

You have exceeded the maximum number of locks allowed on a recordset. This limit is specified by the MaxLocksPerFile setting in your system registry. The default value is 9500, and can be changed either by editing the registry with Regedit.exe or with the **SetOption** method.

Some other factors that may cause an application to reach this threshold include the following:


- amount of available memory
    
- size of rows in the recordset
    
- network operating system restrictions
    
## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](https://learn.microsoft.com/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
