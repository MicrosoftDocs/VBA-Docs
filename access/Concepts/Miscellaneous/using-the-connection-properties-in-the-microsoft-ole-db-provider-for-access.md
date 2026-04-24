---
title: Using the Connection Properties in the Microsoft OLE DB Provider for Access
keywords: acmain11.chm1032169
f1_keywords:
- acmain11.chm1032169
ms.assetid: 7bf8c7d0-9185-d7b2-505a-6ddc449089b9
ms.date: 06/08/2017
ms.localizationpriority: medium
---


# Using the Connection Properties in the Microsoft OLE DB Provider for Access

  

**Applies to:** Access 2013 | Access 2016

To customize the Microsoft® Windows® Registry settings, you can use the connection properties in the Microsoft OLE DB Provider for Access. This is accomplished by referencing a property in the connection object and changing its value. For example, assuming that your connection object is called ADOConnection, the following would yield the same results as going through ADO: 

ADOConnection.Properties("Jet OLEDB:Max Locks Per File") = 20000
The property names are different than the DAO constants and the registry settings. The property names are as follows:
Jet OLEDB:Max Locks Per File
Jet OLEDB:Implicit Commit Sync
Jet OLEDB:Flush Transaction Timeout
Jet OLEDB:Lock Delay
Jet OLEDB:Max Buffer Size
Jet OLEDB:User Commit Sync
Jet OLEDB:Lock Retry
Jet OLEDB:Exclusive Async Delay
Jet OLEDB:Shared Async Delay
Jet OLEDB:Page Timeout
Jet OLEDB:Recycle Long-Valued Pages

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](https://learn.microsoft.com/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]