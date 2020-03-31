---
title: ORDER BY clause <clause> conflicts with GROUP BY clause. (Error 3094)
keywords: jeterr40.chm5003094
f1_keywords:
- jeterr40.chm5003094
ms.prod: access
ms.assetid: 8b6878ef-113c-69e6-5265-3e70e9dd4408
ms.date: 06/08/2019
localization_priority: Normal
---


# ORDER BY clause <clause> conflicts with GROUP BY clause. (Error 3094)

  

**Applies to:** Access 2013 | Access 2016

This error can occur when you create a select query and specify grouped fields after sorted fields that are not grouped. For example, you created an SQL statement with the specified field out of sequence in the ORDER BY clause. Move the specified field to the left of the ungrouped fields, or remove the specified field from the ORDER BY clause.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]