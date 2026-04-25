---
title: Cannot update <field name>; field not updatable. (Error 3113)
keywords: jeterr40.chm5003113
f1_keywords:
- jeterr40.chm5003113
ms.assetid: a86b3fc0-f78f-d9dc-963d-3fbe710a4be9
ms.date: 06/08/2019
ms.localizationpriority: medium
---


# Cannot update \<field name\>; field not updatable. (Error 3113)

  

**Applies to:** Access 2013 | Access 2016

Possible causes:



- The specified field is part of a **TableDef** or dynaset-type **Recordset** object that cannot be updated. For example, this error occurs if you try to update an AutoNumber field.
    
- You executed a query that combines updatable and nonupdatable **TableDef** objects, and you tried to update one of the fields in the query's results (the resulting dynaset-type **Recordset** ).
    
## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]