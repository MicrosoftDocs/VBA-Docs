---
title: Invalid scale for decimal data type. (Error 3701)
keywords: jeterr40.chm5003701
f1_keywords:
- jeterr40.chm5003701
ms.prod: access
ms.assetid: 2c839f39-d3ab-053a-d7b0-1bcde43232d4
ms.date: 06/08/2017
localization_priority: Normal
---


# Invalid scale for decimal data type. (Error 3701)

  

**Applies to:** Access 2013 | Access 2016

The scale of a DECIMAL data type must always be less or equal to the precision. For example, the following SQL statement would return this error: CREATE TABLE foo (foo DECIMAL(10,12));

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]