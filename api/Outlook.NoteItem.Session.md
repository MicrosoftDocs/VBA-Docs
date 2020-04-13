---
title: NoteItem.Session property (Outlook)
keywords: vbaol11.chm1475
f1_keywords:
- vbaol11.chm1475
ms.prod: outlook
api_name:
- Outlook.NoteItem.Session
ms.assetid: 87ebd38c-eec8-7e2c-8516-6ad3053e06cb
ms.date: 06/08/2017
localization_priority: Normal
---


# NoteItem.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [NoteItem](Outlook.NoteItem.md) object.


## Remarks

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[NoteItem Object](Outlook.NoteItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]