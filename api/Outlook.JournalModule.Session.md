---
title: JournalModule.Session property (Outlook)
keywords: vbaol11.chm2864
f1_keywords:
- vbaol11.chm2864
ms.prod: outlook
api_name:
- Outlook.JournalModule.Session
ms.assetid: 416b232d-bed3-fcf5-db47-2946b5a8d244
ms.date: 06/08/2017
localization_priority: Normal
---


# JournalModule.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [JournalModule](Outlook.JournalModule.md) object.


## Remarks

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[JournalModule Object](Outlook.JournalModule.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]