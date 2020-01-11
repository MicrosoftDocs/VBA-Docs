---
title: Invalid PageTimeout setting in the Windows Registry. (Error 3249)
keywords: jeterr40.chm5003249
f1_keywords:
- jeterr40.chm5003249
ms.prod: access
ms.assetid: 62962a57-4e33-ea83-76ee-c20428deda7d
ms.date: 06/08/2017
localization_priority: Normal
---


# Invalid PageTimeout setting in the Windows Registry. (Error 3249)

  

**Applies to:** Access 2013 | Access 2016

There is an invalid  **PageTimeout** setting in the Microsoft Windows Registry.

 To complete this operation


1. Exit your application.
    
2. Start the Registry Editor, and navigate to the  **PageTimeout** value. Depending on which installable ISAM you are trying to use, the invalid entry is in the **HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\14.0\Access Connectivity Engine\Engines\Xbase** key.
    
3. On the  **Edit** menu, click **Modify**.
    
4. Specify a new value in the  **Value data** box.
    
5. Restart your application, and then try the operation again.
    

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]