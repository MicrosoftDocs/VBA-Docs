---
title: The object cannot be replicated because the database is not replicated. (Error 3485)
keywords: jeterr40.chm5003485
f1_keywords:
- jeterr40.chm5003485
ms.assetid: ca11f046-2fa6-6da3-89ba-eacab953a992
ms.date: 06/08/2017
ms.localizationpriority: medium
---


# The object cannot be replicated because the database is not replicated. (Error 3485)

  

**Applies to:** Access 2013 | Access 2016

You cannot replicate an object in a database unless you first replicate the database that contains it. You can replicate the database by:



- Dragging it into the Microsoft Windows Briefcase.
    
- Using DAO programming to set the **Replicable** property to "T" or the **ReplicableBool** property to **True**.
    
- Using Microsoft Access.
    
- Using the Replication Manager.
    

All objects in the database are replicated when the database is replicated, unless the **KeepLocal** property has been set on an object.

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](https://learn.microsoft.com/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]