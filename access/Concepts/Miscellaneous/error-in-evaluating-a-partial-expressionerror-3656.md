---
title: Error in evaluating a partial expression. (Error 3656)
ms.assetid: 4426220f-f086-8bd6-3a61-452e95c0b3da
ms.date: 06/08/2019
ms.localizationpriority: medium
---


# Error in evaluating a partial expression. (Error 3656)

  

**Applies to:** Access 2013 | Access 2016

You have entered an invalid expression in a Boolean filter used to determine which records to use in a partial replica. The replica filter can take the following values:



|Value|Description|
|:-----|:-----|
|A string|A criteria that a record must satisfy in order to appear in the replicated table. The string is similar to an SQL WHERE clause, but you cannot specify subqueries, aggregate functions (such as Count), or user-defined functions within the criteria.|
|True|Replicate all records.|
|False|(Default) Do not replicate any records.|

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]