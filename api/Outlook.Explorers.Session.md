---
title: Explorers.Session property (Outlook)
keywords: vbaol11.chm118
f1_keywords:
- vbaol11.chm118
ms.prod: outlook
api_name:
- Outlook.Explorers.Session
ms.assetid: 51dede9c-3775-2ca9-553e-5bd87ff35ae6
ms.date: 06/08/2017
localization_priority: Normal
---


# Explorers.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents an [Explorers](Outlook.Explorers.md) object.


## Remarks

The  **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[Explorers Object](Outlook.Explorers.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]