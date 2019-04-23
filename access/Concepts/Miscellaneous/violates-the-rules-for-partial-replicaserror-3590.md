---
title: Violates the rules for partial replicas. (Error 3590)
keywords: jeterr40.chm5003590
f1_keywords:
- jeterr40.chm5003590
ms.prod: access
ms.assetid: e8cb495b-cf7d-3a81-f49c-d1c8f837956e
ms.date: 06/08/2017
localization_priority: Normal
---


# Violates the rules for partial replicas. (Error 3590)

  

**Applies to:** Access 2013 | Access 2016

You cannot update a column in a table in a partial replica when another table references that column. Most likely, this is an update RI case, where the related table information does not exist in the partial replica. Make sure you follow all relationships to related tables when defining your Partial Filters.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]