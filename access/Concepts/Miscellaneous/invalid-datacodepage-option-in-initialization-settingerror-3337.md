---
title: Invalid DataCodePage option in initialization setting. (Error 3337)
ms.assetid: 51df967e-82dd-38c3-e413-dfbf728d065d
ms.date: 06/08/2019
ms.localizationpriority: medium
---


# Invalid DataCodePage option in initialization setting. (Error 3337)

  

**Applies to:** Access 2013 | Access 2016

The **DataCodePage** setting for the external data source you are attempting to use is not valid. This setting is in the corresponding **HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\14.0\Access Connectivity Engine\Engines** _&lt;external data source ISAM>_ in the Microsoft Windows Registry.

Valid settings are:


- **OEM** — Data is stored as OEM data; OemToAnsi and AnsiToOem conversions are done.
    
- **ANSI** — Data is stored as ANSI data; OemToAnsi and AnsiToOem conversions are not done.
    

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]