---
title: Invalid Mark setting in the Xbase key of the Windows Registry. (Error 3214)
ms.prod: access
ms.assetid: 3d64dd79-921b-c04d-45b6-52c457199744
ms.date: 06/08/2017
localization_priority: Normal
---


# Invalid Mark setting in the Xbase key of the Windows Registry. (Error 3214)

  

**Applies to:** Access 2013 | Access 2016

There is an invalid  **Mark** setting in the **Xbase** key of the Microsoft Windows Registry.

 To fix the Mark setting


1. Exit your application.
    
2. Start the Registry Editor, navigate to the  **HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\14.0\Access Connectivity Engine\Engines\Xbase** key, and select the **Mark** value.
    
3. On the  **Edit** menu, click **Modify**.
    
4. Correct the  **Mark** data in the **Value data** box.
    
5. Restart your application, and then try the operation again.
    

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]