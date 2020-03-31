---
title: Cannot open this Paradox 4.x or 5.x table because ParadoxNetStyle is set to 3.x in the Windows Registry. (Error 3419)
keywords: jeterr40.chm5003419
f1_keywords:
- jeterr40.chm5003419
ms.prod: access
ms.assetid: 020660b7-10ab-15dc-54f5-091ed149bac4
ms.date: 06/08/2019
localization_priority: Normal
---


# Cannot open this Paradox 4.x or 5.x table because ParadoxNetStyle is set to 3.x in the Windows Registry. (Error 3419)

  

**Applies to:** Access 2013 | Access 2016

The Microsoft Access database engine was unable to open an external Paradox table because of inconsistencies between the version of the table and your  **ParadoxNetStyle** setting in the Microsoft Windows Registry.

To correct the problem


1. Exit your application.
    
2. Start the Registry Editor, navigate to the  **HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\14.0\Access Connectivity Engine\Engines\Paradox** key, and select the **ParadoxNetStyle** value.
    
3. On the  **Edit** menu, click **Modify**.
    
4. Specify the value  **4.x** in the **Value data** box.
    
5. The  **ParadoxNetStyle** setting must be consistent for all users sharing a database.
    
6. Restart your application.
    
7. Try opening the table again.
    

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]