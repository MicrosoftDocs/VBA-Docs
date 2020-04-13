---
title: Pages.Session property (Outlook)
keywords: vbaol11.chm393
f1_keywords:
- vbaol11.chm393
ms.prod: outlook
api_name:
- Outlook.Pages.Session
ms.assetid: cebf5807-8f1f-05f4-e990-35fb00e07f0a
ms.date: 06/08/2017
localization_priority: Normal
---


# Pages.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [Pages](Outlook.Pages.md) object.


## Remarks

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[Pages Object](Outlook.pages(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]