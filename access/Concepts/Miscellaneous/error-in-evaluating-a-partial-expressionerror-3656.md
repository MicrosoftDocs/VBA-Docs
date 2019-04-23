---
title: Error in evaluating a partial expression. (Error 3656)
ms.prod: access
ms.assetid: 4426220f-f086-8bd6-3a61-452e95c0b3da
ms.date: 06/08/2017
localization_priority: Normal
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

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]