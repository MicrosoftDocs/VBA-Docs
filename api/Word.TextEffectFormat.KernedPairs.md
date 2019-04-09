---
title: TextEffectFormat.KernedPairs property (Word)
keywords: vbawd10.chm164561001
f1_keywords:
- vbawd10.chm164561001
ms.prod: word
api_name:
- Word.TextEffectFormat.KernedPairs
ms.assetid: 555d152e-09ff-b151-46c6-9a14ab872a37
ms.date: 06/08/2017
localization_priority: Normal
---


# TextEffectFormat.KernedPairs property (Word)

Indicates that character pairs in a WordArt object have been kerned. Read/write  **MsoTriState**.


## Syntax

_expression_. `KernedPairs`

_expression_ Required. A variable that represents a '[TextEffectFormat](Word.TextEffectFormat.md)' object.


## Example

This example turns on character pair kerning for all WordArt objects in the active document.


```vb
Sub Kerned() 
 With ActiveDocument.Range(1, ActiveDocument.Shapes.Count).ShapeRange 
 If .Type = msoTextEffect Then 
 .TextEffect.KernedPairs = True 
 End If 
 End With 
End Sub
```


## See also


[TextEffectFormat Object](Word.TextEffectFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]