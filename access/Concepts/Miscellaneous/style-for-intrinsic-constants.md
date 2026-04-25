---
title: Style for Intrinsic Constants
keywords: vbaac10.chm4050
f1_keywords:
- vbaac10.chm4050
ms.assetid: 6f301835-307b-d0b8-be24-c0fa728cc115
ms.date: 06/08/2017
ms.localizationpriority: medium
---


# Style for Intrinsic Constants

  

**Applies to:** Access 2013 | Access 2016

In Microsoft Access, all intrinsic constants are contained in type libraries and are visible in the Object Browser. Microsoft Access includes type libraries for Microsoft Access, ActiveX Data Objects (ADO), Data Access Objects (DAO), and Visual Basic. Each of these type libraries includes intrinsic constants.

Additionally, intrinsic constants in Microsoft Access are a mix of lowercase and uppercase, and parts of the constant are concatenated rather than separated by underscores. For example, the constant A_NORMAL in versions 1. _x_ and 2.0 is now **acNormal**.
Intrinsic constants in databases created with previous versions of Microsoft Access won't automatically be converted to the new constant format, but old constants will continue to work without errors. However, it's recommended that you use the new format when writing new code.

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]