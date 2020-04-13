---
title: ItemProperties.Session property (Outlook)
keywords: vbaol11.chm533
f1_keywords:
- vbaol11.chm533
ms.prod: outlook
api_name:
- Outlook.ItemProperties.Session
ms.assetid: 5dde3402-b791-e0f7-e4fe-10bb85e5284a
ms.date: 06/08/2017
localization_priority: Normal
---


# ItemProperties.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents an [ItemProperties](Outlook.ItemProperties.md) object.


## Remarks

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[ItemProperties Object](Outlook.ItemProperties.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]