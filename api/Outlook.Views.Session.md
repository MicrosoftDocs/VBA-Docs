---
title: Views.Session property (Outlook)
keywords: vbaol11.chm543
f1_keywords:
- vbaol11.chm543
ms.prod: outlook
api_name:
- Outlook.Views.Session
ms.assetid: 677d7b97-b138-3506-7b45-26d091f9ba6e
ms.date: 06/08/2017
localization_priority: Normal
---


# Views.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [Views](Outlook.Views.md) object.


## Remarks

The  **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[Views Object](Outlook.Views.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]