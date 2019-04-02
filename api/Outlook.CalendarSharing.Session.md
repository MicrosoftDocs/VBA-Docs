---
title: CalendarSharing.Session property (Outlook)
keywords: vbaol11.chm2409
f1_keywords:
- vbaol11.chm2409
ms.prod: outlook
api_name:
- Outlook.CalendarSharing.Session
ms.assetid: 5a20615c-7639-8009-2fdf-86b28ba591ba
ms.date: 06/08/2017
localization_priority: Normal
---


# CalendarSharing.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [CalendarSharing](Outlook.CalendarSharing.md) object.


## Remarks

The  **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[CalendarSharing Object](Outlook.CalendarSharing.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]