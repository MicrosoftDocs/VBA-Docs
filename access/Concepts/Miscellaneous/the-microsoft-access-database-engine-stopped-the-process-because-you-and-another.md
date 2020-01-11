---
title: The Microsoft Access database engine stopped the process because you and another user are attempting to change the same data at the same time. (Error 3197)
keywords: jeterr40.chm5003197
f1_keywords:
- jeterr40.chm5003197
ms.prod: access
ms.assetid: 3ea30548-166c-2cfc-5014-6d624a75294e
ms.date: 06/08/2017
localization_priority: Normal
---


# The Microsoft Access database engine stopped the process because you and another user are attempting to change the same data at the same time. (Error 3197)

  

**Applies to:** Access 2013 | Access 2016

This error can occur in a multiuser environment.

Another user has changed the data you are trying to update. This error can occur when multiple users open a table or create a  **Recordset** and use optimistic locking. Between the time you used the **Edit** method and the **Update** method, another user changed the same data.
To overwrite the other user's changes with your own, execute the  **Update** method again.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
