---
title: UndoRecord.EndCustomRecord method (Word)
keywords: vbawd10.chm56098818
f1_keywords:
- vbawd10.chm56098818
ms.prod: word
api_name:
- Word.UndoRecord.EndCustomRecord
ms.assetid: af11d231-f799-d592-2bc5-de08030b41e4
ms.date: 06/08/2017
localization_priority: Normal
---


# UndoRecord.EndCustomRecord method (Word)

Completes the creation of a custom undo record.


## Syntax

_expression_. `EndCustomRecord`

_expression_ A variable that represents an '[UndoRecord](Word.UndoRecord.md)' object.


## Remarks

You use the [UndoRecord.StartCustomRecord](Word.UndoRecord.StartCustomRecord.md) to initiate the creation of a custom undo record. To complete the creation of a custom undo record, you use the **EndCustomRecord** method.


## Example

The following code example creates a custom undo record.


```vb
Sub TestUndo() 
Dim objUndo As UndoRecord 
 
Set objUndo = Application.UndoRecord 
objUndo.StartCustomRecord ("My Custom Undo") 
    'Add some actions here 
objUndo.EndCustomRecord 
     
End Sub
```


## See also


[UndoRecord Object](Word.UndoRecord.md)



[Working with the UndoRecord Object](../word/Concepts/Working-with-Word/working-with-the-undorecord-object.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]