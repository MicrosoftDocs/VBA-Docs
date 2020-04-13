---
title: Row.Session property (Outlook)
keywords: vbaol11.chm2241
f1_keywords:
- vbaol11.chm2241
ms.prod: outlook
api_name:
- Outlook.Row.Session
ms.assetid: a9773e62-0091-50b4-f64c-dab4217035cc
ms.date: 06/08/2017
localization_priority: Normal
---


# Row.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [Row](Outlook.Row.md) object.


## Remarks

The **Session** property and the **[Application.GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[Row Object](Outlook.Row.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]