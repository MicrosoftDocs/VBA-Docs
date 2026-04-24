---
title: Invalid NetworkAccess setting in the Windows Registry. (Error 3248)
ms.assetid: c61fb87b-47eb-9644-f4af-5c5676e8c2b8
ms.date: 06/08/2019
ms.localizationpriority: medium
---


# Invalid NetworkAccess setting in the Windows Registry. (Error 3248)

  

**Applies to:** Access 2013 | Access 2016

There is an invalid **NetworkAccess** setting in the Microsoft Windows Registry.

 To complete this operation


1. Exit your application.
    
2. Start the Registry Editor, and navigate to the **NetworkAccess** value. Depending on which installable ISAM you are trying to use, the invalid entry is in the **HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\14.0\Access Connectivity Engine\Engines\Xbase** key.
    
3. On the **Edit** menu, click **Modify**.
    
4. Specify either **On** or **Off** in the **Value data** box.
    
5. Restart your application, and then try the operation again.
    

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](https://learn.microsoft.com/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]