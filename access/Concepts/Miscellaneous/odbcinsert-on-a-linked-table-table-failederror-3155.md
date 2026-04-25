---
title: ODBC - insert on a linked table <table> failed. (Error 3155)
keywords: jeterr40.chm5003155
f1_keywords:
- jeterr40.chm5003155
ms.assetid: 96d34c06-a906-d215-ff4f-8be70da27e6a
ms.date: 06/08/2019
ms.localizationpriority: medium
---


# ODBC - insert on a linked table <table> failed. (Error 3155)

  

**Applies to:** Access 2013 | Access 2016

Using an ODBC connection, you tried to insert data into an ODBC database; that insertion could not be completed.

Possible causes:


- The insertion would have caused a rule violation.
    
- The ODBC database is read-only, or you don't have permission to insert data in that database. Resolve the read-only condition, or see your system administrator or the person who created the database to obtain the necessary permissions.
    
- The ODBC database is on a network drive, and you are not connected to the network. Make sure the network is available, and then try the operation again.
    

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
