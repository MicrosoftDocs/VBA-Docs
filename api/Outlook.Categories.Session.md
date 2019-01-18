---
title: Categories.Session Property (Outlook)
keywords: vbaol11.chm2433
f1_keywords:
- vbaol11.chm2433
ms.prod: outlook
api_name:
- Outlook.Categories.Session
ms.assetid: f810b08c-bf94-d4f6-563f-b0329af37f74
ms.date: 06/08/2017
localization_priority: Normal
---


# Categories.Session Property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_. `Session`

_expression_ A variable that represents a [Categories](./Outlook.Categories.md) object.


## Remarks

The  **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[Categories Object](Outlook.Categories.md)

