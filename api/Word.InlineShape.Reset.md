---
title: InlineShape.Reset method (Word)
keywords: vbawd10.chm162005093
f1_keywords:
- vbawd10.chm162005093
ms.prod: word
api_name:
- Word.InlineShape.Reset
ms.assetid: c7c7c01a-7c62-7d2f-24e6-d20c02c8dbb3
ms.date: 06/08/2017
localization_priority: Normal
---


# InlineShape.Reset method (Word)

Removes changes that were made to an inline shape.


## Syntax

_expression_. `Reset`

_expression_ Required. A variable that represents an '[InlineShape](Word.InlineShape.md)' object.


## Example

This example inserts a picture as an inline shape, changes the brightness, and then resets the picture to its original brightness.


```vb
Set aInLine = ActiveDocument.InlineShapes.AddPicture _ 
 (FileName:="C:\Windows\Bubbles.bmp", Range:=Selection.Range) 
aInLine.PictureFormat.Brightness = 0.5 
MsgBox "Changing brightness back" 
aInLine.Reset
```


## See also


[InlineShape Object](Word.InlineShape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]