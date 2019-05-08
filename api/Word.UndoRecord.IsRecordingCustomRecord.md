---
title: UndoRecord.IsRecordingCustomRecord property (Word)
keywords: vbawd10.chm56098819
f1_keywords:
- vbawd10.chm56098819
ms.prod: word
api_name:
- Word.UndoRecord.IsRecordingCustomRecord
ms.assetid: 08693e04-4a76-f7ab-9671-cdad35ac87ea
ms.date: 06/08/2017
localization_priority: Normal
---


# UndoRecord.IsRecordingCustomRecord property (Word)

Returns a  **Boolean** that specifies whether a custom undo action is being recorded. Read-only.


## Syntax

_expression_. `IsRecordingCustomRecord`

_expression_ A variable that represents an **[UndoRecord](Word.UndoRecord.md)** object.


## Example

The following code example displays whether a custom undo action is currently being recorded.


```vb
Dim objUndo as UndoRecord 
Set objUndo = Application.UndoRecord 
 
If objUndo.IsRecordingCustomRecord = False Then 
objUndo.StartCustomRecord ("My Custom Undo") 
End If 
'Custom undo actions here 
objUndo.EndCustomRecord 


```


## See also


[UndoRecord Object](Word.UndoRecord.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]