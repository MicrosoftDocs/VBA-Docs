---
title: CalendarModule.Session property (Outlook)
keywords: vbaol11.chm2824
f1_keywords:
- vbaol11.chm2824
ms.prod: outlook
api_name:
- Outlook.CalendarModule.Session
ms.assetid: df23c975-9ac9-4ed9-0369-dce6b59e518a
ms.date: 06/08/2017
localization_priority: Normal
---


# CalendarModule.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [CalendarModule](Outlook.CalendarModule.md) object.


## Remarks

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[CalendarModule Object](Outlook.CalendarModule.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]