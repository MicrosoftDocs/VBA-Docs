---
title: SharingItem.Session property (Outlook)
keywords: vbaol11.chm595
f1_keywords:
- vbaol11.chm595
ms.prod: outlook
api_name:
- Outlook.SharingItem.Session
ms.assetid: 0563b8e1-6b5b-2c1f-9cc7-3c69ccb80346
ms.date: 06/08/2017
localization_priority: Normal
---


# SharingItem.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [SharingItem](Outlook.SharingItem.md) object.


## Remarks

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[SharingItem Object](Outlook.SharingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]