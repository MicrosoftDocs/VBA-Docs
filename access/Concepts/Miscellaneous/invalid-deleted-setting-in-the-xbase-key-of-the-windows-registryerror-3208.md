---
title: Invalid Deleted setting in the Xbase key of the Windows Registry. (Error 3208)
ms.assetid: f28f55fc-2332-d008-8d8f-013175c22f5b
ms.date: 06/08/2019
ms.localizationpriority: medium
---


# Invalid Deleted setting in the Xbase key of the Windows Registry. (Error 3208)

  

**Applies to:** Access 2013 | Access 2016

There is an invalid **Deleted** setting in the **Xbase** key of the Microsoft Windows Registry.

 To fix the Deleted setting


1. Exit your application.
    
2. Start the Registry Editor, navigate to the **HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\14.0\Access Connectivity Engine\Engines\Xbase** key, and select the **Deleted** value.
    
3. On the **Edit** menu, click **Modify**.
    
4. Correct the **Deleted** data in the **Value data** box.
    
5. Restart your application, and then try the operation again.
    

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]