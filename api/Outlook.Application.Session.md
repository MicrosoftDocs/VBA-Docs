---
title: Application.Session property (Outlook)
keywords: vbaol11.chm707
f1_keywords:
- vbaol11.chm707
ms.prod: outlook
api_name:
- Outlook.Application.Session
ms.assetid: 720b2849-fe01-afb3-363c-f3bf0cd7d872
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents an **[Application](Outlook.Application.md)** object.


## Remarks

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[Application Object](Outlook.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
