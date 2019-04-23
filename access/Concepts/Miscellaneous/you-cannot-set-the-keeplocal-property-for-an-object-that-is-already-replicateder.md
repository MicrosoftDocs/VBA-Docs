---
title: You cannot set the KeepLocal property for an object that is already replicated. (Error 3457)
keywords: jeterr40.chm5003457
f1_keywords:
- jeterr40.chm5003457
ms.prod: access
ms.assetid: 916ea4af-3190-99f4-901d-76b7754efa6a
ms.date: 06/08/2017
localization_priority: Normal
---


# You cannot set the KeepLocal property for an object that is already replicated. (Error 3457)

  

**Applies to:** Access 2013 | Access 2016

The  **KeepLocal** property cannot be set on a replicated object. Setting a local object's **KeepLocal** property after the database has been replicated has no effect on the object. If you want to keep an object from being replicated to the other replicas in the set, set the object's **Replicable** property to "F".

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]