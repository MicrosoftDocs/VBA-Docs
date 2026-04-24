---
title: Object is invalid or no longer set. (Error 3420)
keywords: jeterr40.chm5003420
f1_keywords:
- jeterr40.chm5003420
ms.assetid: 5744c5e1-1cf7-52eb-6ac3-a35044f2f6d6
ms.date: 06/08/2019
ms.localizationpriority: medium
---


# Object is invalid or no longer set. (Error 3420)

  

**Applies to:** Access 2013 | Access 2016

You are attempting to reference an object that is no longer valid or has not been set.

Possible causes:


- The object has been closed.
    
- The object has been orphaned (the parent object has been closed or deleted).
    
- The object is out of scope.
    
- The object library is not registered in the Microsoft Windows Registry.
    
- You are trying to reference a method or property of the collection, but you have not assigned it to a variable first. For example, to reference the **Name** property, use the following:
    
```vb
  Dim dbsPublish As Database 
Set dbsPublish = OpenDatabase("BIBLIO.mdb")
dbname = dbsPublish.Name

```

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](https://learn.microsoft.com/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
