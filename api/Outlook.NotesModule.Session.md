---
title: NotesModule.Session property (Outlook)
keywords: vbaol11.chm2874
f1_keywords:
- vbaol11.chm2874
ms.prod: outlook
api_name:
- Outlook.NotesModule.Session
ms.assetid: 066a38fa-3b6a-ea03-9bee-23ec95c89388
ms.date: 06/08/2017
localization_priority: Normal
---


# NotesModule.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [NotesModule](Outlook.NotesModule.md) object.


## Remarks

The  **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[NotesModule Object](Outlook.NotesModule.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]