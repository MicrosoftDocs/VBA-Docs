---
title: Field <name> is based on an expression and cannot be edited. (Error 3327)
ms.prod: access
ms.assetid: 7d7c1e1f-645e-b111-60c3-666640d8bde1
ms.date: 06/08/2017
localization_priority: Normal
---


# Field <name> is based on an expression and cannot be edited. (Error 3327)

  

**Applies to:** Access 2013 | Access 2016

For example, if a stored query or view with a column made up of an expression was created, you would not be able to update that column. The following would return this error: CREATE VIEW VCustomer AS SELECT (FirstName & LastName) AS Test FROM Customer followed by UPDATE Test FROM VCustomer

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]