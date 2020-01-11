---
title: Cannot open <tablename>. Another user has the table open using a different network control file or locking style. (Error 3418)
keywords: jeterr40.chm5003418
f1_keywords:
- jeterr40.chm5003418
ms.prod: access
ms.assetid: 3f3c4b8b-0749-61f1-f8dd-635f836cf335
ms.date: 06/08/2017
localization_priority: Normal
---


# Cannot open <tablename>. Another user has the table open using a different network control file or locking style. (Error 3418)

  

**Applies to:** Access 2013 | Access 2016

The Microsoft Access database engine cannot open an external Paradox table because of inconsistencies between your initialization settings and those of another user who currently has the table open. The  **ParadoxNetPath** and the **ParadoxNetStyle** settings in the **HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\14.0\Access Connectivity Engine\Engines\Paradox** key of the Microsoft Windows Registry must be consistent for all users sharing a database. Make sure your initialization settings match those of all other users sharing the database, and then try opening the table again.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]