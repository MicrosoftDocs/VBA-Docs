---
title: Could not build key. (Error 3250)
keywords: jeterr40.chm5003250
f1_keywords:
- jeterr40.chm5003250
ms.prod: access
ms.assetid: c00debc3-c39d-6c58-6206-f0210a6e1ea4
ms.date: 06/08/2017
localization_priority: Normal
---


# Could not build key. (Error 3250)

  

**Applies to:** Access 2013 | Access 2016

When building a primary index, the Microsoft Access database engine could not build a primary key. Make sure the key fields are named properly and that there are no duplicate records based on this key.

This error can occur when you use the  **Seek** method and pass it a value for a field that is not part of the index. For example, suppose you want to use the **Seek** method on a **Recordset** whose current index uses the LastName field of the underlying table, and you write it this way:

    rstEmployees.Seek "=", "Smith", "Joe"

The Microsoft Access database engine will try to construct a primary key from two fields, but because one field is indexed, the attempt will fail and this error results.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]