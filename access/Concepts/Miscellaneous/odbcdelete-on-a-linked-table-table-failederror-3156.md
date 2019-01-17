---
title: ODBC - delete on a linked table <table> failed. (Error 3156)
keywords: jeterr40.chm5003156
f1_keywords:
- jeterr40.chm5003156
ms.prod: access
ms.assetid: fd432abe-d5ee-66c4-90fa-cfd697d45d62
ms.date: 06/08/2017
localization_priority: Normal
---


# ODBC - delete on a linked table <table> failed. (Error 3156)

  

**Applies to:** Access 2013 | Access 2016

Using an ODBC connection, you tried to delete data from an ODBC database; the deletion could not be completed.

Possible causes:


- The deletion would have caused a rule violation.
    
- The ODBC database is read-only, or you do not have permission to delete data in that database. Resolve the read-only condition, or see your system administrator or the person who created the database to obtain the necessary permissions.
    
- The ODBC database is on a network drive and the network is not connected. Make sure the network is available, and then try the operation again.
    

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]