---
title: SELECT INTO on a remote database tried to produce too many fields. (Error 3185)
keywords: jeterr40.chm5003185
f1_keywords:
- jeterr40.chm5003185
ms.prod: access
ms.assetid: 82f94c46-9a79-7a8f-dc69-86bf9a272f52
ms.date: 06/08/2017
localization_priority: Normal
---


# SELECT INTO on a remote database tried to produce too many fields. (Error 3185)

  

**Applies to:** Access 2013 | Access 2016

The Microsoft Access database engine supports up to 255 fields per table, but the other database supports fewer fields. A SELECT...INTO statement created a table with more fields than the remote database can support. Reduce the number of columns produced by the SELECT INTO statement, and then try the operation again.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]