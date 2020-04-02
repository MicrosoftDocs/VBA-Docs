---
title: Invalid precision for decimal data type. (Error 3700)
ms.prod: access
ms.assetid: 74d04222-2ed3-1552-4b56-f205633efd7f
ms.date: 06/08/2019
localization_priority: Normal
---


# Invalid precision for decimal data type. (Error 3700)

  

**Applies to:** Access 2013 | Access 2016

The precision and scale for a DECIMAL data type must be between 0 and 28. For example, the following SQL statement would return this error: CREATE TABLE foo (foo DECIMAL(38,0));

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]