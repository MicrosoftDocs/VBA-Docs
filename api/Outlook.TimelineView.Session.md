---
title: TimelineView.Session property (Outlook)
keywords: vbaol11.chm2644
f1_keywords:
- vbaol11.chm2644
ms.prod: outlook
api_name:
- Outlook.TimelineView.Session
ms.assetid: 9d85749d-c254-c294-112f-d0343a2f01a9
ms.date: 06/08/2017
localization_priority: Normal
---


# TimelineView.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [TimelineView](Outlook.TimelineView.md) object.


## Remarks

The  **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[TimelineView Object](Outlook.TimelineView.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]