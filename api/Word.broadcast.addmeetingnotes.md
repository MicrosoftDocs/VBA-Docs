---
title: Broadcast.AddMeetingNotes method (Word)
keywords: vbawd10.chm36438121
f1_keywords:
- vbawd10.chm36438121
ms.prod: word
ms.assetid: e13a52fd-d0a4-bc32-2d0a-f01f9218bfa2
ms.date: 06/08/2017
localization_priority: Normal
---


# Broadcast.AddMeetingNotes method (Word)

Adds shared meeting notes for the specified broadcast that are accessible to attendees who use either Microsoft OneNote 2013 rich client or web app.


## Syntax

_expression_.**AddMeetingNotes** (_notesUrl_, _notesWacUrl_)

_expression_ A variable that represents a **[Broadcast](Word.broadcast.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _notesUrl_|Required|**String**|Specifies the URL where the shared meeting notes are stored, for attendees using the Microsoft OneNote 2013 rich client.|
| _notesWacUrl_|Required|**String**|Specifies the URL where the shared meeting notes are stored, for attendees using the Microsoft OneNote 2013 web access client.|

## Return value

**VOID**


## Remarks

If you fail to pass a string for either of the two parameters, the **AddMeetingNotes** method returns an Invalid Parameter error. If for any reason the method call fails, Word returns a generic broadcast error.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]