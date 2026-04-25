---
title: Invalid operation. (Error 3219)
keywords: jeterr40.chm5003219
f1_keywords:
- jeterr40.chm5003219
ms.assetid: ab31a5dd-0979-2a03-3816-ef62ac370cae
ms.date: 06/08/2019
ms.localizationpriority: medium
---


# Invalid operation. (Error 3219)

  

**Applies to:** Access 2013 | Access 2016

Possible causes:



- You tried to write to a read-only property. See the Help topic for the property to determine whether it is read/write.
    
- You tried to use a method or property on a type of **Recordset** object that the method or property does not apply to. See the Recordset object summary topic to determine which methods and properties apply to a given type of **Recordset** object.
    
- You tried to append a property to a **Properties** collection of an object that does not support user-defined properties.
    
- You tried to use the **Update** method on a read-only **Recordset** object.
    

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]