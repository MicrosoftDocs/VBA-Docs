---
title: The record cannot be deleted or changed because table <name> includes related records. (Error 3200)
keywords: jeterr40.chm5003200
f1_keywords:
- jeterr40.chm5003200
ms.prod: access
ms.assetid: e3171406-6a42-5932-35f4-b0a4db616f3a
ms.date: 06/08/2017
localization_priority: Normal
---


# The record cannot be deleted or changed because table <name> includes related records. (Error 3200)

  

**Applies to:** Access 2013 | Access 2016

You tried to perform an operation that would have violated referential integrity rules for related tables. For example, this error occurs if you try to delete or change a record in the "one" table in a one-to-many relationship when there are related records in the "many" table.

If you want to delete or change the record, first delete the related records from the "many" table.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
