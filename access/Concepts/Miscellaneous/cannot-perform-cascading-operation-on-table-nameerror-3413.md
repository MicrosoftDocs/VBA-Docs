---
title: Cannot perform cascading operation on table <name>. (Error 3413)
keywords: jeterr40.chm5003413
f1_keywords:
- jeterr40.chm5003413
ms.prod: access
ms.assetid: 0a060fbf-86a5-24b4-8a46-2f2ea90ea4ab
ms.date: 06/08/2019
localization_priority: Normal
---


# Cannot perform cascading operation on table <name>. (Error 3413)

  

**Applies to:** Access 2013 | Access 2016

This error occurs when referential integrity is defined with one of the following actions: CASCADE DELETE, UPDATE or NULL. The error occurs when making a change to a table that contains a referenced primary key and one or more of the foreign key tables is opened in a mode that prevents the foreign key tables from being opened in a shared read/write mode.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]