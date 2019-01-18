---
title: Cannot change field <name>. (Error 3720)
keywords: jeterr40.chm5003720
f1_keywords:
- jeterr40.chm5003720
ms.prod: access
ms.assetid: c3062dac-0f7c-be1b-d9ee-48cd178d0241
ms.date: 06/08/2017
localization_priority: Normal
---


# Cannot change field <name>. (Error 3720)

  

**Applies to:** Access 2013 | Access 2016

This error occurs when trying to change a name or definition of a primary key that is referenced by other foreign keys. To change the definition or name of the columns participating in the primary key, you must first remove all foreign key references to the primary key. This typically happens when using the ALTER TABLE ALTER COLUMN statement.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]