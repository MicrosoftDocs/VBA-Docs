---
title: In operator without (). (Error 2429)
keywords: jeterr40.chm5002429
f1_keywords:
- jeterr40.chm5002429
ms.assetid: 40f2356c-f891-1d90-17be-ace51c989357
ms.date: 06/08/2019
ms.localizationpriority: medium
---


# In operator without (). (Error 2429)

  

**Applies to:** Access 2013 | Access 2016

When coding an SQL statement that includes the **In** operator, you must surround the list of items to test with parentheses. For example, to see if a value is one of a set of values, you could use the following code in the WHERE clause of an SQL query:




```vb
WHERE Region In ('TX', 'CA', 'WA')

```

This code tests to see if the Region field contains any of the above abbreviations, which represent Texas, California, and Washington.

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](https://learn.microsoft.com/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]