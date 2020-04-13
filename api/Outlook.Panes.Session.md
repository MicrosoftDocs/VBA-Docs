---
title: Panes.Session property (Outlook)
keywords: vbaol11.chm76
f1_keywords:
- vbaol11.chm76
ms.prod: outlook
api_name:
- Outlook.Panes.Session
ms.assetid: 3f0eeae2-e02e-d7f1-70de-6c9d869756d9
ms.date: 06/08/2017
localization_priority: Normal
---


# Panes.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [Panes](Outlook.Panes.md) object.


## Remarks

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[Panes Object](Outlook.Panes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]