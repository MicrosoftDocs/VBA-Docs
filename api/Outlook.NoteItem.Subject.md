---
title: NoteItem.Subject property (Outlook)
keywords: vbaol11.chm1489
f1_keywords:
- vbaol11.chm1489
ms.prod: outlook
api_name:
- Outlook.NoteItem.Subject
ms.assetid: 17c4d857-e548-e0fb-475d-8764bcd0f17d
ms.date: 06/08/2017
localization_priority: Normal
---


# NoteItem.Subject property (Outlook)

Returns or sets a  **String** indicating the subject for the Outlook item. Read-only.


## Syntax

_expression_. `Subject`

_expression_ A variable that represents a [NoteItem](Outlook.NoteItem.md) object.


## Remarks

The  **Subject** property is a **String** that is calculated from the body text of the note.

This property corresponds to the MAPI property  **PidTagSubject**. The **Subject** property is the default property for Outlook items.


## See also


[NoteItem Object](Outlook.NoteItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]