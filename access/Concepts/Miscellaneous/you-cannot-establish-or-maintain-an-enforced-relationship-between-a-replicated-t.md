---
title: You cannot establish or maintain an enforced relationship between a replicated table and a local table. (Error 3453)
keywords: jeterr40.chm5003453
f1_keywords:
- jeterr40.chm5003453
ms.assetid: 1bd3124e-452f-4cd7-8c71-dbc3267e63a8
ms.date: 06/08/2017
ms.localizationpriority: medium
---


# You cannot establish or maintain an enforced relationship between a replicated table and a local table. (Error 3453)

  

**Applies to:** Access 2013 | Access 2016

You are attempting to establish or maintain an enforced relationship between a replicated table and a non-replicated table. Replication does not allow an enforced relationship between:



- A replicated table and a local table.
    
- Two local tables that you are making replicable.
    
- Two tables with different **KeepLocal** property settings.
    

Delete the relationship between the two tables before proceeding.

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](https://learn.microsoft.com/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]