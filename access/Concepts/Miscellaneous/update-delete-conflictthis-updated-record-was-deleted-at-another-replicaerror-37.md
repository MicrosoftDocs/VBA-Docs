---
title: Update/delete conflict - This updated record was deleted at another replica. (Error 3736)
keywords: jeterr40.chm5003736
f1_keywords:
- jeterr40.chm5003736
ms.prod: access
ms.assetid: d8e66115-9a71-72b1-137b-61305057fb00
ms.date: 06/08/2017
localization_priority: Normal
---


# Update/delete conflict - This updated record was deleted at another replica. (Error 3736)

  

**Applies to:** Access 2013 | Access 2016

When a record is deleted at one replica, but updated at another replica, the deleted record always wins in the conflict that occurs when the two replicas synchronize. The updated record is logged in the conflict table. To reverse the initial resolution of the conflict, reinsert the conflict record. To accept the current resolution, delete the conflict record.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]