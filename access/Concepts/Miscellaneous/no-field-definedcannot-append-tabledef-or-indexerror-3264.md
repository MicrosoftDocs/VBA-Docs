---
title: No field defined -- cannot append TableDef or Index. (Error 3264)
ms.prod: access
ms.assetid: 18353c1b-c3c7-9f41-eb2a-87d732d2127a
ms.date: 06/08/2019
localization_priority: Normal
---


# No field defined -- cannot append TableDef or Index. (Error 3264)

  

**Applies to:** Access 2013 | Access 2016

You cannot append a **TableDef** until you define one or more fields. Use the **CreateField** method to create fields, append them to the **Fields** collection of your **TableDef** object, and then append the **TableDef** object to the **TableDefs** collection.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]