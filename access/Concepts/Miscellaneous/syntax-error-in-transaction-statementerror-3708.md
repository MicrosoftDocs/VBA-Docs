---
title: Syntax error in Transaction statement. (Error 3708)
keywords: jeterr40.chm5003708
f1_keywords:
- jeterr40.chm5003708
ms.prod: access
ms.assetid: 3533b96d-05fd-1e16-6bb3-090af470c46a
ms.date: 06/08/2017
localization_priority: Normal
---


# Syntax error in Transaction statement. (Error 3708)

  

**Applies to:** Access 2013 | Access 2016

When using transactions exclusively through the Microsoft Access database engine (not through an object model like DAO or ADO) the following syntax must be used: BEGIN/ROLLBACK TRANSACTION, WORK or nothing. If any other characters follow BEGIN or ROLLBACK, then this error will occur.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]