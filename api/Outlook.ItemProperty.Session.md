---
title: ItemProperty.Session property (Outlook)
keywords: vbaol11.chm520
f1_keywords:
- vbaol11.chm520
ms.prod: outlook
api_name:
- Outlook.ItemProperty.Session
ms.assetid: f33cfcd0-f86b-d0cd-7d35-a21644bc5c42
ms.date: 06/08/2017
localization_priority: Normal
---


# ItemProperty.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents an [ItemProperty](Outlook.ItemProperty.md) object.


## Remarks

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[ItemProperty Object](Outlook.ItemProperty.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]