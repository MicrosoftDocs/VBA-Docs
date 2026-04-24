---
title: Wait For Post Processing property
keywords: vbaac10.chm5991
f1_keywords:
- vbaac10.chm5991
ms.assetid: b747ff33-3e84-480c-bcd8-3b5e7d0e063d
ms.date: 06/08/2017
ms.localizationpriority: medium
---


# Wait For Post Processing property

  

**Applies to:** Access 2013 | Access 2016

Use the **Wait For Post Processing** property to specify that the form waits until processing of any operations (for example, running a macro) triggered by a user change to form data is complete before proceeding with the next operation.


## Setting

The **Wait For Post Processing** property uses the following settings.



|Setting|Description|
|:-----|:-----|
|**Yes**|Wait until processing of any operations triggered by a user change to form data is complete before proceeding with the next operation.|
|**No**|(Default) Does not wait until processing of any operations triggered by a user change to form data is complete before proceeding with the next operation.|

## Remarks

This property is designed to work with Access 2010 web databases only. When this property is set to **Yes**, if a user changes data in a form that then triggers a data macro, the form will wait for the macro to finish before proceeding.

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]