---
title: You cannot add or change a record because a related record is required in table <name>. (Error 3201)
keywords: jeterr40.chm5003201
f1_keywords:
- jeterr40.chm5003201
ms.prod: access
ms.assetid: 371ebb17-1809-8076-9bde-aea1df33ef74
ms.date: 06/08/2017
localization_priority: Priority
---


# You cannot add or change a record because a related record is required in table <name>. (Error 3201)

  

**Applies to:** Access 2013 | Access 2016

You tried to perform an operation that would have violated referential integrity rules for related tables. For example, this error occurs if you try to change or insert a record in the "many" table in a one-to-many relationship, and that record does not have a related record in the table on the "one" side.

If you want to add or change the record, first add a record to the "one" table that contains the same value for the matching field.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
