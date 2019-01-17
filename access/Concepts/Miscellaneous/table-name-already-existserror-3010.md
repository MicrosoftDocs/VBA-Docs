---
title: Table <name> already exists. (Error 3010)
keywords: jeterr40.chm5003010
f1_keywords:
- jeterr40.chm5003010
ms.prod: access
ms.assetid: eaf04a4d-15cb-b27a-d6e3-8f89e88d4143
ms.date: 06/08/2017
localization_priority: Normal
---


# Table <name> already exists. (Error 3010)

  

**Applies to:** Access 2013 | Access 2016

You tried to create or rename a table with a name that already exists in this database. Choose another name, and then try the operation again.

In a multiuser database, this error can also occur if you delete a table, another user creates a table with the same name, and then you try to roll back the deletion of your table. To restore your table, the other user must first delete or rename the new table before you try the rollback operation again.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]