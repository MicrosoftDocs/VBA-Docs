---
title: ViewFont.Session property (Outlook)
keywords: vbaol11.chm2693
f1_keywords:
- vbaol11.chm2693
ms.prod: outlook
api_name:
- Outlook.ViewFont.Session
ms.assetid: 8f126189-3bec-6eee-1e62-b178738d361b
ms.date: 06/08/2017
localization_priority: Normal
---


# ViewFont.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [ViewFont](Outlook.ViewFont.md) object.


## Remarks

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[ViewFont Object](Outlook.ViewFont.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]