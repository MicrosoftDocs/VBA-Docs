---
title: UndoRecord object (Word)
keywords: vbawd10.chm856
f1_keywords:
- vbawd10.chm856
ms.prod: word
api_name:
- Word.UndoRecord
ms.assetid: 77bf9801-e940-e661-6bbe-20a8714d5dbd
ms.date: 06/08/2017
localization_priority: Normal
---


# UndoRecord object (Word)

Provides an entry point into the undo stack.


## Remarks

Use the **UndoRecord** object to create and modify custom undo records in the Word undo stack.


## Example

The following code example instantiates an **UndoRecord** object.


```vb
Dim objUndo As UndoRecord 
Set objUndo = Application.UndoRecord
```


## See also


[Working With the UndoRecord Object](../word/Concepts/Working-with-Word/working-with-the-undorecord-object.md)
[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]