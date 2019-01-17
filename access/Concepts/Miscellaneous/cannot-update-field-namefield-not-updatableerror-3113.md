---
title: Cannot update <field name>; field not updatable. (Error 3113)
keywords: jeterr40.chm5003113
f1_keywords:
- jeterr40.chm5003113
ms.prod: access
ms.assetid: a86b3fc0-f78f-d9dc-963d-3fbe710a4be9
ms.date: 06/08/2017
localization_priority: Normal
---


# Cannot update <field name>; field not updatable. (Error 3113)

  

**Applies to:** Access 2013 | Access 2016

Possible causes:



- The specified field is part of a  **TableDef** or dynaset-type **Recordset** object that cannot be updated. For example, this error occurs if you try to update an AutoNumber field.
    
- You executed a query that combines updatable and nonupdatable  **TableDef** objects, and you tried to update one of the fields in the query's results (the resulting dynaset-type **Recordset** ).
    
## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]