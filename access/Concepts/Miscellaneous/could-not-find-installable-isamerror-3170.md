---
title: Could not find installable ISAM. (Error 3170)
keywords: jeterr40.chm5003170
f1_keywords:
- jeterr40.chm5003170
ms.prod: access
ms.assetid: 1a97fb83-4732-0f8f-9fb0-d5a11236797c
ms.date: 06/08/2017
localization_priority: Normal
---


# Could not find installable ISAM. (Error 3170)

  

**Applies to:** Access 2013 | Access 2016

The DLL for an installable ISAM file could not be found. This file is required for linking external tables (other than ODBC or Microsoft Access database engine tables). The locations for all ISAM drivers are maintained in the Microsoft Windows Registry. These entries are created automatically when you install your application. If you change the location of these drivers, you need to correct your application Setup program to reflect this change and make the correct entries in the Registry.

Possible causes:


- An entry in the Registry is not valid. For example, this error occurs if you are using a Paradox external database, and the Paradox entry points to a nonexistent directory or driver. Exit the application, correct the Windows Registry, and try the operation again.
    
- One of the entries in the Registry points to a network drive and that network is not connected. Make sure the network is available, and then try the operation again.
    

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]