---
title: UndoRecord.StartCustomRecord method (Word)
keywords: vbawd10.chm56098817
f1_keywords:
- vbawd10.chm56098817
ms.prod: word
api_name:
- Word.UndoRecord.StartCustomRecord
ms.assetid: cd8d4337-4bbc-1943-6e0a-bc764861e886
ms.date: 06/08/2017
localization_priority: Normal
---


# UndoRecord.StartCustomRecord method (Word)

Initiates the creation of a custom undo record.


## Syntax

_expression_. `StartCustomRecord`( `_Name_` )

_expression_ A variable that represents an '[UndoRecord](Word.UndoRecord.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Optional| **String**|Specifies the name of the custom undo record. This string is limited to 64 characters. If a longer string is supplied, the string is truncated to 64 characters. 

> [!NOTE] 
> If this parameter is omitted or is an empty string, Word uses the name of the first command executed as the name of the undo record.

|

## Remarks

 **StartCustomRecord** begins the creation of a custom undo record, which records all actions done to the application while it is active under a record defined by _Name_.


## Example


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