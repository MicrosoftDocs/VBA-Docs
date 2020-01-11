---
title: The number of aliases specified shall be the same as number of output columns (Error 3731)
keywords: jeterr40.chm5003731
f1_keywords:
- jeterr40.chm5003731
ms.prod: access
ms.assetid: 884d1f65-60d7-66f3-f404-d7b0b996c46a
ms.date: 06/08/2017
localization_priority: Normal
---


# The number of aliases specified shall be the same as number of output columns (Error 3731)

  

**Applies to:** Access 2013 | Access 2016

This error occurs when trying to create a view through SQL DDL. The error occurs when a different number of correlation names or aliases are defined from what is in the SELECT statement. For example, the following syntax would generate this error: CREATE VIEW foo (col1, col2) AS SELECT col1 FROM table1.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]