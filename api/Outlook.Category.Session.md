---
title: Category.Session property (Outlook)
keywords: vbaol11.chm2424
f1_keywords:
- vbaol11.chm2424
ms.prod: outlook
api_name:
- Outlook.Category.Session
ms.assetid: e942f0c1-930f-fe1f-0b57-fe4b2894ee74
ms.date: 06/08/2017
localization_priority: Normal
---


# Category.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [Category](Outlook.Category.md) object.


## Remarks

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[Category Object](Outlook.Category.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]