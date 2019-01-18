---
title: Invalid Century setting in the Xbase key of the Windows Registry. (Error 3227)
ms.prod: access
ms.assetid: eb75b1ac-b7d7-d151-7ea0-4ecb7e265a51
ms.date: 06/08/2017
localization_priority: Normal
---


# Invalid Century setting in the Xbase key of the Windows Registry. (Error 3227)

  

**Applies to:** Access 2013 | Access 2016

There is an invalid  **Century** setting in the **Xbase** key of the Microsoft Windows Registry.

To fix the Century setting


1. Exit your application.
    
2. Start the Registry Editor, navigate to the  **HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\14.0\Access Connectivity Engine\Engines\Xbase** key, and select the **Century** value.
    
3. On the  **Edit** menu, click **Modify**.
    
4. Correct the  **Century** data in the **Value data** box.
    
5. Restart your application, and then try the operation again.
    

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]