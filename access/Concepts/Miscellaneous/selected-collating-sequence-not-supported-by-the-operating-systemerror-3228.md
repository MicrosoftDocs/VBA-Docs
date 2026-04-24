---
title: Selected collating sequence not supported by the operating system. (Error 3228)
keywords: jeterr40.chm5003228
f1_keywords:
- jeterr40.chm5003228
ms.assetid: 39eae021-6584-478d-cb75-ce5131341dba
ms.date: 06/08/2017
ms.localizationpriority: medium
---


# Selected collating sequence not supported by the operating system. (Error 3228)

  

**Applies to:** Access 2013 | Access 2016

There is an invalid **CollatingSequence** setting in the **Paradox** key of the Microsoft Windows Registry.

 To fix the CollatingSequence setting


1. Exit your application.
    
2. Start the Registry Editor, navigate to the **HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\14.0\Access Connectivity Engine\Engines\Paradox** key, and select the **CollatingSequence** value.
    
3. On the **Edit** menu, click **Modify**.
    
4. Specify either **ascii** or **International** in the **Value data** box.
    
5. Restart your application, and then try the operation again.
    

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](https://learn.microsoft.com/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]