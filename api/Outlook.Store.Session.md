---
title: Store.Session property (Outlook)
keywords: vbaol11.chm798
f1_keywords:
- vbaol11.chm798
ms.prod: outlook
api_name:
- Outlook.Store.Session
ms.assetid: 90dc9dc2-41c5-6448-4f42-98d8e4a6f948
ms.date: 06/08/2017
localization_priority: Normal
---


# Store.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [Store](Outlook.Store.md) object.


## Remarks

The  **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[Store Object](Outlook.Store.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]