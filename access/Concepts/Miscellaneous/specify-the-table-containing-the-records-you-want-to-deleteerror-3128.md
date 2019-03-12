---
title: Specify the table containing the records you want to delete. (Error 3128)
keywords: jeterr40.chm5003128
f1_keywords:
- jeterr40.chm5003128
ms.prod: access
ms.assetid: f6c49cba-5b9c-775c-625a-6d1e79c8adf0
ms.date: 06/08/2017
localization_priority: Normal
---


# Specify the table containing the records you want to delete. (Error 3128)

  

**Applies to:** Access 2013 | Access 2016

You tried to execute a delete query but the query does not specify the name of the table containing the records you want to delete.

Possible cause:


- You did not type an asterisk for each table in the ALL, DISTINCT, DISTINCTROW predicates. Instead, you typed field names (for example,  `Customers.Address` instead of `Customers.*`).
    

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
