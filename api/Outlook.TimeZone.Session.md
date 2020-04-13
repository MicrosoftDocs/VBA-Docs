---
title: TimeZone.Session property (Outlook)
keywords: vbaol11.chm3280
f1_keywords:
- vbaol11.chm3280
ms.prod: outlook
api_name:
- Outlook.TimeZone.Session
ms.assetid: 8b696765-dcc5-3af2-a861-a14c9c0bf7e8
ms.date: 06/08/2017
localization_priority: Normal
---


# TimeZone.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [TimeZone](Outlook.TimeZone.md) object.


## Remarks

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[TimeZone Object](Outlook.TimeZone.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]