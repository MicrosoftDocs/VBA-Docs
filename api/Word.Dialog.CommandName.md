---
title: Dialog.CommandName property (Word)
keywords: vbawd10.chm163085574
f1_keywords:
- vbawd10.chm163085574
ms.prod: word
api_name:
- Word.Dialog.CommandName
ms.assetid: 5bd7a993-b40e-57ca-65c7-260efcea488b
ms.date: 06/08/2017
localization_priority: Normal
---


# Dialog.CommandName property (Word)

Returns the name of the procedure that displays the specified built-in dialog box. Read-only  **String**.


## Syntax

_expression_. `CommandName`

_expression_ A variable that represents a '[Dialog](Word.Dialog.md)' object.


## Remarks

For more information about working with built-in Word dialog boxes, see [Displaying Built-in Word Dialog Boxes](../word/Concepts/Customizing-Word/displaying-built-in-word-dialog-boxes.md).


## Example

This example displays the name of the procedure that displays the  **Save As** dialog box (**File** menu): **FileSaveAs**.


```vb
MsgBox Dialogs(wdDialogFileSaveAs).CommandName
```


## See also


[Dialog Object](Word.Dialog.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]