---
title: InlineShape.HasSmartArt property (Word)
keywords: vbawd10.chm162005147
f1_keywords:
- vbawd10.chm162005147
ms.prod: word
api_name:
- Word.InlineShape.HasSmartArt
ms.assetid: fd53f446-d0f9-5d67-7369-b2fdd241da4e
ms.date: 06/08/2017
localization_priority: Normal
---


# InlineShape.HasSmartArt property (Word)

Returns  **True** if there is a SmartArt diagram present on the shape. Read-only.


## Syntax

 _expression_. `HasSmartArt`

 _expression_ A variable that represents an '[InlineShape](Word.InlineShape.md)' object.


## Example

The following code example displays whether or not the first inline shape in the active document contains SmartArt.


```vb
Dim myInlineShape As InlineShape 
 
Set myInlineShape = ActiveDocument.InlineShapes(1) 
 
If myInlineShape.HasSmartArt Then 
 MsgBox "The first shape contains SmartArt." 
Else 
 MsgBox "The first shape contains no SmartArt." 
End If
```


## See also


[InlineShape Object](Word.InlineShape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]