---
title: System.Cursor property (Word)
keywords: vbawd10.chm154468368
f1_keywords:
- vbawd10.chm154468368
ms.prod: word
api_name:
- Word.System.Cursor
ms.assetid: f4acf757-920f-f389-948e-e2a142d451b0
ms.date: 06/08/2017
localization_priority: Normal
---


# System.Cursor property (Word)

Returns or sets the state (shape) of the pointer. Can be one of the following  **WdCursorType** constants: **wdCursorIBeam**, **wdCursorNormal**, **wdCursorNorthwestArrow**, or **wdCursorWait**. Read/write **Long**.


## Syntax

_expression_. `Cursor`

_expression_ A variable that represents a '[System](Word.System.md)' object.


## Example

This example prints a message on the status bar and changes the pointer to a busy pointer.


```vb
Dim intWait As Integer 
 
StatusBar = "Please wait..." 
 
For intWait = 1 To 1000 
 System.Cursor = wdCursorWait 
Next intWait 
 
StatusBar = "Task completed" 
System.Cursor = wdCursorNormal
```


## See also


[System Object](Word.System.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]