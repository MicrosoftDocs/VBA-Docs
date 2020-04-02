---
title: Paradox index is not primary. (Error 3288)
keywords: jeterr40.chm5003288
f1_keywords:
- jeterr40.chm5003288
ms.prod: access
ms.assetid: 5866732f-b013-a777-a3ed-46a9d2ce43ea
ms.date: 06/08/2019
localization_priority: Normal
---


# Paradox index is not primary. (Error 3288)

  

**Applies to:** Access 2013 | Access 2016

You are attempting to create an index on a Paradox table for which no other indexes are defined. The first index created on a Paradox table must be a primary index.

Possible solutions:


- Redefine the index as the primary index.
    
- If the index you are attempting to create is not the primary index, create the primary index first.
    

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]