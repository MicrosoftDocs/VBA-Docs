---
title: The Microsoft Access database engine could not execute the SQL statement because it contains a field that has an invalid data type. (Error 3169)
keywords: jeterr40.chm5003169
f1_keywords:
- jeterr40.chm5003169
ms.assetid: 0d1c107b-4bf9-e389-c2da-cb1ea14fa65e
ms.date: 06/08/2017
ms.localizationpriority: medium
---


# The Microsoft Access database engine could not execute the SQL statement because it contains a field that has an invalid data type. (Error 3169)

  

**Applies to:** Access 2013 | Access 2016

You tried to execute an SQL statement that contains a field that has an invalid field data type.

Possible causes:


- You included a Memo or OLE Object field in an expression where it is not allowed.
    
- You included a numeric aggregate function, such as **Sum** or **StDev**, that tried to perform a calculation on a Text field. Choose a different aggregate function.
    

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](https://learn.microsoft.com/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]